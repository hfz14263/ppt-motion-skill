#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regression for the privacy audit's secret rule and its file set.

为什么需要它
------------
privacy_audit 是"0 条违规"的来源，而 0 有两种可能：真的干净，或判据写错了
（CODE_RULES §十.4.1）。2026-10-12 复查发现这条密钥规则两个方向都错：

  * 误报 —— 把"秘密的**引用来源**"当成"秘密本身"：`os.environ["GH_TOKEN"]`、
    `$(git credential fill)` 这些**正确做法**全被报出来；一个总在误报的规则
    会训练人忽略它。
  * 漏报（更要紧）—— `\\bsecret\\b` 的 `\\b` 在 `_` 后面不成立，所以
    `client_secret` / `access_token` / `db_password` —— 最常见的命名形式，
    原规则**一个都看不见**。

修完必须把判据钉住，否则下次"顺手简化正则"会静默回到坏状态。
两条清单代表两个方向：
  * CATCH —— 字面量赋值，必须报（旧规则 8 条里只报 3 条）；
  * ALLOW —— 引用 / 占位符，必须放过（旧规则 12 条里错报 7 条）。

再钉住一次文件集：.gitignore 的构建产物（out/）不进**文本**扫描 ——
旧实现把 out/ 里一个 PDF 的压缩字节读成"邮箱地址"，那是最典型的
"判据自己造出来的违规"；而**容器**扫描保留 out/（那是要递给人的 deck）、
只排除 `.workbuddy/`（每跑一次探测/测试就再生的本机残留，会把信号淹掉）。
"""
import os
import shutil
import subprocess
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import privacy_audit as pa                                   # noqa: E402

# 字面量赋值：必须报出来（全部是编造的假值，本文件因此进 SKIP_FILES）
CATCH = [
    'client_secret = "AbCdEf0123456789"',
    'access_token = "abc123XYZ789def456"',
    'db_password = "hunter2hunter2"',
    'MY_SECRET_KEY = "AbCdEf0123456789"',
    'GH_TOKEN = "AbCdEf0123456789"',
    'API_KEY=AbCdEf0123456789AbCd',
    'secret: "AbCdEf0123456789"',
    'password: correcthorsebatterystaple',
]

# 引用来源 / 占位符：必须放过（这正是"取密码的正确姿势"）
ALLOW_LINES = [
    'token = os.environ["GH_TOKEN"]',
    'GH_TOKEN="$TOKEN" REPO_DIR="$(pwd)"',
    "TOKEN=$(printf 'protocol=https\\nhost=github.com\\n\\n' | git credential fill \\",
    'GH_TOKEN=<pat> REPO_DIR=<本地仓库> python scripts/push_via_api.py',
    'password = getpass()',
    'api_key = os.getenv("API_KEY")',
    'token = self._get_token()',
    'api_key = config["api_key"]',
    'token: <your-token-here>',
    'SECRET = os.environ.get("SECRET")',
    'secret: ${SECRET_FROM_ENV}',
    'token_type = "Bearer"',
]

fails = []


def check(label, ok):
    print("   [%s] %s" % ("ok" if ok else "FAIL", label))
    if not ok:
        fails.append(label)


def secret_findings(lines):
    """Run the real scan_text over the sample lines; return the hit line numbers."""
    hits = pa.scan_text("\n".join(lines))
    return set(ln for ln, label, _snip in hits if label == 'possible secret')


def scope_checks():
    """文本扫描的候选集 = 将随提交出门的文件；.gitignore 的产物不在其中。"""
    # 「没有 git」才 skip；git 在而判据没生效 = FAIL —— 静默跳过就是
    # CODE_RULES §十.4.1 说的那种"分不清的 0"。
    if subprocess.run(["git", "--version"], capture_output=True).returncode != 0:
        print("   [skip] 这台机器没有 git —— 文件集判据未验证")
        return
    tmp = tempfile.mkdtemp(prefix="pa_scope_")
    try:
        subprocess.run(["git", "init", "-q", tmp], capture_output=True)

        def w(relpath, s):
            with open(os.path.join(tmp, relpath), "w", encoding="utf-8") as fh:
                fh.write(s)

        w(".gitignore", "out/\n")
        os.makedirs(os.path.join(tmp, "out"), exist_ok=True)
        w(os.path.join("out", "trap.pdf"), "x someone@example.com x\n")
        w("keep.txt", "y someone@example.com y\n")
        files = pa.worktree_files(tmp)
        rel = ([os.path.relpath(f, tmp).replace(os.sep, "/") for f in files]
               if files else [])
        check("git 在，worktree_files 必须生效（不是 None）", files is not None)
        check("被 .gitignore 的 out/trap.pdf 不进候选集", "out/trap.pdf" not in rel)
        check("未跟踪但会被提交的 keep.txt 在候选集里", "keep.txt" in rel)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def container_scope_checks():
    """容器扫描的树：含 out/（要递给人的 deck），不含 .workbuddy/（本机探测残留）。"""
    tmp = tempfile.mkdtemp(prefix="pa_cont_")
    try:
        os.makedirs(os.path.join(tmp, ".workbuddy"))
        os.makedirs(os.path.join(tmp, "out"))
        for rel in (os.path.join(".workbuddy", "probe.pptx"),
                    os.path.join("out", "deck.pptx"),
                    "plain.pptx"):
            open(os.path.join(tmp, rel), "wb").write(b"")
        got = set(os.path.relpath(p, tmp).replace(os.sep, "/")
                  for p in pa.iter_container_files(tmp))
        check("容器扫描跳过 .workbuddy/（本机探测残留，会淹掉信号）",
              ".workbuddy/probe.pptx" not in got)
        check("容器扫描包含 out/（要递给人的 deck，作者名最要紧）",
              "out/deck.pptx" in got)
        check("容器扫描包含普通文件", "plain.pptx" in got)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    print("== privacy_audit：密钥规则正/负样本 ==")
    caught = secret_findings(CATCH)
    for i, line in enumerate(CATCH, 1):
        check("报出字面量  %s" % line[:52], i in caught)
    for line in ALLOW_LINES:
        check("放过引用/占位符  %s" % line[:46], not secret_findings([line]))
    print()
    print("== privacy_audit：文本扫描的文件集 ==")
    scope_checks()
    print()
    print("== privacy_audit：容器扫描的树 ==")
    container_scope_checks()
    print()
    if fails:
        print("-- FAIL --")
        for f in fails:
            print("   " + f)
        return 1
    print("privacy_audit ok  （CATCH %d 条全报出，ALLOW %d 条全放过，文件集 3 + 容器树 3 判据）"
          % (len(CATCH), len(ALLOW_LINES)))
    return 0


if __name__ == "__main__":
    sys.exit(main())

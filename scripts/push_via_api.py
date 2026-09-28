#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""当 github.com 被代理挡住、但 api.github.com 通时，用 Git Data API 推送。

背景见 reference/com-pitfalls.md §45。典型症状：

    git push
    → CONNECT tunnel failed, response 502
    但 curl -x <proxy> https://api.github.com 通（200）

这不是凭证问题，是通道问题 —— 本机代理不放行 github.com:443，而 git push
走的就是它。既然 API 通，就用 API 手工造对象：
    blob -> tree -> commit -> PATCH ref
只处理"本地领先远端若干个 commit"这一种情况。

用法：
    GH_TOKEN=<pat> REPO_DIR=<本地仓库> python scripts/push_via_api.py

⚠️ 关键点（踩过一次，见 §45.3）：
    core.autocrlf=true 时，工作区是 CRLF、对象库是 LF。**必须上传
    git cat-file 出来的字节**，读工作区文件会让 GitHub 算出另一个 sha，
    建 tree 时报 "is not a valid blob"。

⚠️ 校验（见 §45.4）：API 造的 commit sha 必然与本地不同，别拿 sha 判断
    成败；要比 **tree 哈希**（相等即内容逐字节一致）。
"""
import base64
import json
import os
import subprocess
import sys
import urllib.error
import urllib.request

REPO = os.environ.get("REPO", "hfz14263/ppt-motion-skill")
API = "https://api.github.com"
PROXY = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy") or ""


def api(method, path, payload=None):
    token = os.environ["GH_TOKEN"]
    url = API + path
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header("Authorization", "Bearer " + token)
    req.add_header("Accept", "application/vnd.github+json")
    req.add_header("User-Agent", "api-push")
    if data:
        req.add_header("Content-Type", "application/json")
    handlers = [urllib.request.ProxyHandler(
        {"https": PROXY, "http": PROXY})] if PROXY else []
    op = urllib.request.build_opener(*handlers)
    try:
        with op.open(req, timeout=60) as r:
            body = r.read().decode("utf-8")
            return r.status, (json.loads(body) if body else {})
    except urllib.error.HTTPError as e:
        return e.code, json.loads(e.read().decode("utf-8") or "{}")


def git(*args):
    return subprocess.check_output(("git",) + args,
                                   cwd=os.environ["REPO_DIR"]).decode().strip()


# GitHub 返回的新 commit sha 只在远端存在，本地 ls-tree 看不到。
# 维护 远端sha -> 本地sha 的映射，用于拿"父 commit 的 tree"做差集。
SHA_MAP = {}


def _local_of(sha):
    return SHA_MAP.get(sha, sha)


def _tree_files(sha):
    """{path: (mode, blob_sha)} —— 用于和父 commit 做差集。"""
    out = {}
    raw = git("ls-tree", "-r", "-z", sha)
    for item in raw.split("\0"):
        if not item:
            continue
        meta, path = item.split("\t")
        mode, typ, bsha = meta.split()
        out[path] = (mode, bsha)
    return out


def main():
    repo_dir = os.environ["REPO_DIR"]
    # 远端当前 main
    st, ref = api("GET", "/repos/%s/git/ref/heads/main" % REPO)
    assert st == 200, (st, ref)
    remote_sha = ref["object"]["sha"]
    # 本地要推的 commit（HEAD）
    local_sha = git("rev-parse", "HEAD")
    if remote_sha == local_sha:
        print("远端已是最新，无需推送")
        return 0

    # 远端 -> HEAD 之间的提交，从旧到新
    chain = git("rev-list", "--reverse", "%s..%s" % (remote_sha, local_sha)).split()
    if not chain:
        print("没有需要推送的提交")
        return 0
    print("远端 %s -> 本地 %s，共 %d 个提交" % (remote_sha[:7], local_sha[:7], len(chain)))

    parent = remote_sha
    for sha in chain:
        msg = git("log", "-1", "--format=%B", sha)
        author = git("log", "-1", "--format=%an <%ae>", sha)
        date = git("log", "-1", "--format=%aI", sha)

        # 该 commit 的完整文件清单（含 mode / path / blob sha 复用）
        tab = _tree_files(sha)
        entries = [(p, m, b) for p, (m, b) in tab.items()]
        # 只上传父树里没有的 blob —— 用「父 commit 的 tree 里是否存在同 path 同 sha」
        # 判断，而不是 GET /git/blobs/<sha>：后者只认仓库已有对象，会误判。
        have = set()
        ptab = _tree_files(_local_of(parent)) if parent else {}
        for path, (mode, bsha) in ptab.items():
            have.add(path + "|" + bsha)
        n_new = 0
        for path, mode, bsha in entries:
            if (path + "|" + bsha) in have:
                continue
            # ⚠️ 必须用 git 对象里的字节，不能用工作区文件。
            # 本仓库 core.autocrlf=true：工作区是 CRLF、对象库是 LF，
            # 直接读工作区会让 GitHub 算出另一个 sha，tree 就会报
            # "tree.sha ... is not a valid blob"。
            content = subprocess.check_output(
                ("git", "cat-file", "blob", bsha), cwd=repo_dir)
            st, res = api("POST", "/repos/%s/git/blobs" % REPO,
                          {"content": base64.b64encode(content).decode(),
                           "encoding": "base64"})
            assert st in (200, 201), (path, st, res)
            assert res["sha"] == bsha, (
                "%s: GitHub 算出 %s，git 期望 %s" % (path, res["sha"], bsha))
            n_new += 1
        print("  %s  上传 blob %d / 共 %d 文件" % (sha[:7], n_new, len(entries)))

        # 组装 tree（Git Data API 允许一次性给完整路径）
        tree = [{"path": p, "mode": mode, "type": "blob", "sha": bsha}
                for p, mode, bsha in entries]
        st, tres = api("POST", "/repos/%s/git/trees" % REPO, {"tree": tree})
        assert st in (200, 201), (st, tres)
        st, cres = api("POST", "/repos/%s/git/commits" % REPO,
                       {"message": msg, "tree": tres["sha"], "parents": [parent],
                        "author": {"name": author.split(" <")[0],
                                   "email": author.split(" <")[1].rstrip(">"),
                                   "date": date}})
        assert st in (200, 201), (st, cres)
        parent = cres["sha"]
        SHA_MAP[parent] = sha
        print("  -> 新 commit %s（本地 %s）" % (parent[:7], sha[:7]))

    st, res = api("PATCH", "/repos/%s/git/refs/heads/main" % REPO,
                  {"sha": parent, "force": False})
    assert st == 200, (st, res)
    print("远端 main 已更新到 %s" % parent[:7])
    return 0


if __name__ == "__main__":
    sys.exit(main())

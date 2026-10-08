#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""测试公共 helper —— 只放**三份测试都要用**的东西。

**为什么有这一个文件**：原先22 个测试组全部堆在一个文件里（1633 行），
加第 23 组时你得读完 1600 行才知道该加在哪。拆成三份之后，
`check()` / `_load_json()` / `ROOT` 这些会被重复三遍 —— 所以提到这里。

**这个文件不许长大**。它只许放**跨主题共用**的东西；
只被一份测试用到的 helper 就放在那份测试里。
判据很简单：**另一个测试文件会不会 import 它？**不会 → 搬回去。

模块内容表
------------------------------------------------
   check()        判定并记账（PASS/FAIL 汇总的唯一来源）
   _load_json()   读 JSON —— 专治 PowerShell 写的 BOM
   HERE / ROOT    路径基准，两个测试文件都用
   REF            transition_reference.json 的位置
   LEGACY_SPECS   motion.py 曾手写的 13 个 spec + none
   run_all()      多份测试的汇总入口
   main_of()      把被拆出去的 main() 包成统一签名
"""
import io
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
if os.path.join(ROOT, "scripts") not in sys.path:
    sys.path.insert(0, os.path.join(ROOT, "scripts"))

REF = os.path.join(ROOT, "scripts", "transition_reference.json")

# motion.py 曾经手写的 13 个 spec + none。改名就是 breaking change --
# 用户 spec 里写的 type: push 必须在任意版本里产出同样的东西。
LEGACY_SPECS = ("fade", "smoothfade", "fadeblack", "push", "pushleft", "wipe",
                "cover", "split", "zoom", "dissolve", "strips", "pull",
                "randombar", "none")


def load_json(path):
    """读 JSON —— 必须用 utf-8-sig。

    这些文件有的是 PowerShell 写的，`Set-Content -Encoding UTF8` 会带 BOM，
    用普通 utf-8 读会直接炸。这是本项目最常踩的一个环境坑。
    """
    return json.load(io.open(path, encoding="utf-8-sig"))


def read_text(path):
    """读文本。同样要认 BOM。"""
    with io.open(path, encoding="utf-8-sig") as f:
        return f.read()


def make_checker(title):
    """造一组独立的判定器，返回 (fails, check)。

    **为什么每份测试各造一组**，而不是共用一个全局 fails：
    共用全局会让「A 测试文件的失败混进B 的输出」，
    定位失败时得先想「这是谁的」。分开记账，输出才自解释。
    """
    fails = []

    def check(name, ok, detail=""):
        print("  %-58s %s" % (name, "ok" if ok else "FAIL " + detail))
        if not ok:
            fails.append(name)

    print("== %s ==" % title)
    return fails, check


def report(title, fails):
    """统一收尾：打印失败明细，返回退出码。"""
    print()
    if fails:
        print("%s FAILED (%d):" % (title, len(fails)))
        for f in fails:
            print("  - %s" % f)
        return 1
    print("%s PASSED" % title)
    return 0
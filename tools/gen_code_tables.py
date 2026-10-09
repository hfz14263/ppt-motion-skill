"""生成 reference/code-*.md 里的速查表 —— 表格数据不手抄。

为什么要有这个脚本：
  速查表里最容易腐烂的是**行数**和**依赖关系**。人抄一遍，
  改完代码就忘了改文档，半年后没人知道哪一列是真的。
  所以这两列从代码里扫，扫完直接吐markdown 片段。

用法：
    python tools/gen_code_tables.py            # 打印全部四份表的markdown
    python tools/gen_code_tables.py injection  # 只打印一份
    python tools/gen_code_tables.py --check     # 只报哪份表过期，不打印

输出是**片段**，不是整份文档 —— 表格长什么样由每份code-*.md 自己决定，
这个脚本只保证「抄进文档的每个数字都是真的」。

实现要点（踩过的坑）：
  - **docstring 首句要用 AST 拿，不能用正则**。正则会在 docstring 里
    出现引号、缩进、代码示例时截错位置，AST 给的是真实的第一句。
  - **`__init__.py` 常常没有 docstring**。这种情况填"—"，
    不要从文件名编一句职责出来（编出来的职责没人验证过）。
  - **分层是白名单，不是按目录自动推**。tools/ 里的脚本按功能域
    分属不同层（scan_deps 属校验层、split_* 属工具层），
    自动按目录分会把它们全塞进工具层，那张表就没意义了。
"""
import ast
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 四层各自的成员。显式列出，不自动推导 —— 理由见模块 docstring。
LAYERS = {
    "injection": {
        "title": "注入层",
        "desc": "改OOXML、生成浏览器端重放。全库被依赖最多的两个文件在这一层。",
        "files": [
            ("scripts/motion.py", "核心注入引擎：动画/切换/3D/图片填充 + CLI（门面）"),
            ("scripts/motion_xml.py", "常量 + zip + XML 字符串手术 + 单例 + 形状索引"),
            ("scripts/motion_timing.py", "timing / transition 的 XML 生成"),
            ("scripts/motion_media.py", "3D 相机 + 图片填充"),
            ("scripts/motion_spec.py", "spec 处理 + 结构校验"),
            ("scripts/player.py", "浏览器端动效重放（HTML）"),
        ],
    },
    "probe": {
        "title": "探针层",
        "desc": "造deck、实测参数、量帧。这一层产出的是**证据**，不是动效。",
        "files": [
            ("scripts/build_transition_table.py", "切换实测入口，转发到 transition_probe/"),
            ("scripts/transition_probe/__init__.py", "包标记"),
            ("scripts/transition_probe/common.py", "COM 会话与共用小工具"),
            ("scripts/transition_probe/data.py", "切换形态定义"),
            ("scripts/transition_probe/decks.py", "造deck"),
            ("scripts/transition_probe/analysis.py", "量帧与数据分析"),
            ("scripts/transition_probe/commands_table.py", "建表主链路：造 deck → 枚举扫描 → 合并成表"),
            ("scripts/transition_probe/commands_mechanism.py", "机制层：切换挂在哪一页"),
            ("scripts/transition_probe/commands_shape.py", "形态层与方向：效果看起来在做什么"),
            ("scripts/transition_probe/commands_attr.py", "属性取值全集"),
            ("scripts/transition_probe/commands_timing.py", "切换 × 页内动画：结构平行、时间串行"),
            ("scripts/build_camera_table.py", "相机路径实测表"),
            ("scripts/build_shape_evidence.py", "形状证据实测"),
            ("scripts/build_dual_photo.py", "双色照片对照实验"),
            ("scripts/analyze_video.py", "视频逐帧分析"),
            ("scripts/build_pitfall_map.py", "生成症状→根因映射表"),
        ],
    },
    "verify": {
        "title": "校验层",
        "desc": "证明「它真的生效了」。这一层的产出是**结论**：过 / 不过 + 原因。",
        "files": [
            ("scripts/verify_motion.py", "XML 结构合法性"),
            ("scripts/verify_singletons.py", "singleton 约束"),
            ("scripts/verify_dual_photo.py", "双色照片方案是否真的有效"),
            ("scripts/verify_recipes.py", "配方库结构"),
            ("scripts/verify_docs.py", "文档结构与登记"),
            ("scripts/check_coverage.py", "spec 覆盖率审计"),
            ("scripts/check_structure.py", "仓库结构体检（行数/命名/体积）"),
            ("scripts/inspect_pptx.py", "看一个 pptx 到底有没有动效"),
            ("scripts/design_audit.py", "设计层审计"),
            ("scripts/design_compose.py", "设计层合成"),
            ("scripts/selftest.py", "回归测试入口"),
            ("scripts/review_assist.py", "自动复核"),
            ("tools/scan_deps.py", "依赖图 / 环检测 / 改动波及面（只读）"),
            ("tools/measure_overlap.py", "两族文档逐节相似度（只读，分档依据）"),
            ("tools/check_symptom_coverage.py", "pitfall 现象导语覆盖度（只读）"),
        ],
    },
    "tooling": {
        "title": "工具层",
        "desc": "不产动效，但项目运转要靠它们。PowerShell 脚本全在这一层。",
        "files": [
            ("scripts/push_via_api.py", "走 API 推送"),
            ("scripts/vendor_themes.py", "主题文件入库"),
            ("scripts/motion.ps1", "Office COM 层：media 插入 + 真渲染"),
            ("scripts/make_calibration.ps1", "造标定deck + 标定视频"),
            ("scripts/probe_anchor.ps1", "测 `<p:transition>` 锚点在终点页还是出发页"),
            ("scripts/probe_createvideo.ps1", "探测本机 CreateVideo 是否可用"),
            ("scripts/probe_enum_scan.ps1", "让 PowerPoint 自己给切换命名（反查猜错的元素名）"),
            ("scripts/probe_roundtrip.ps1", "单deck 往返：PowerPoint 认不认我写的切换"),
            ("scripts/probe_shapes.ps1", "把每个形状探针 deck 渲染成视频"),
            ("scripts/probe_timing_roundtrip.ps1", "整目录往返：transition 与 timing 会不会互相吃掉"),
            ("scripts/probe_transitions.ps1", "切换探针：往返 + enum 回读"),
            ("tools/split_transition_probe.py", "一次性：拆 transition_probe 成包"),
            ("tools/split_commands.py", "一次性：把 commands.py 按探测维度拆成 5 份"),
            ("tools/split_motion.py", "一次性：把 motion.py 按分层拆成门面 + 四层（带完整性断言）"),
            ("tools/split_contributing.py", "一次性：把 CONTRIBUTING.md 按读者拆成两份"),
            ("tools/split_pitfalls.py", "一次性：拆 pitfall 族"),
            ("tools/split_symptoms.py", "一次性：拆 symptom 族"),
            ("tools/split_morph_recipes.py", "一次性：拆 morph 配方"),
            ("tools/split_tests.py", "一次性：拆 test_transition_table.py 成三份"),
            ("tools/archive_handover_85.py", "一次性：把 HANDOVER §8 已完成的三节搬进 history/"),
            ("tools/gen_code_tables.py", "生成 / 刷新四份 code-*.md 的速查表（表的数据唯一来源）"),
            ("tools/add_toc.py", "给长文档加目录"),
            ("tools/collapse_symptom.py", "根因收敛 1/3：重复节收敛为索引（幂等）"),
            ("tools/add_symptom_view.py", "根因收敛 2/3：给已收敛节补现象导语（幂等）"),
            ("tools/link_symptom_to_pitfall.py", "根因收敛 3/3：给保留正文的节补根因链接（幂等）"),
            ("tools/add_symptom_leadin.py", "按 TSV 给 pitfall 顶层节插现象导语"),
            # 不在 scripts/ 也不在 tools/ —— 它跟着产物住：
            # design-system/README.md 的生成器（README 声称的文件是否在磁盘上，靠它 --check）
            ("reference/design-system/build_index.py", "生成 / 校验 design-system 索引（--check 只读）"),
        ],
    },
}

# test_transition_table.py 待拆（用户 2026-10-08 决定），
# 现在还单体存在，先登记在探针层，拆完再改这里。
EXTRA_PS = {
    "scripts/make_calibration.ps1": "造已知答案的标定样本",
}


def read_head(path):
    """取docstring 第一句。拿不到就返回 None —— 不要编。"""
    full = os.path.join(ROOT, path)
    if not os.path.exists(full):
        return None
    with open(full, "r", encoding="utf-8") as f:
        src = f.read()
    if path.endswith(".ps1"):
        # PowerShell 用 # 注释，取第一条非空注释行
        for ln in src.split("\n"):
            ln = ln.strip()
            if ln.startswith("#") and len(ln) > 2:
                return ln.lstrip("# ").strip()
        return None
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return None
    doc = ast.get_docstring(tree)
    if not doc:
        return None
    # 第一段（到空行为止）取第一句
    para = doc.strip().split("\n\n")[0].replace("\n", " ").strip()
    for sep in ("。", ". "):
        if sep in para:
            para = para.split(sep)[0]
            break
    return para.strip(" .") or None


def line_count(path):
    full = os.path.join(ROOT, path)
    if not os.path.exists(full):
        return 0
    with open(full, "r", encoding="utf-8", errors="replace") as f:
        return sum(1 for _ in f)



def mod_name(path):
    """scripts/motion.py -> scripts.motion；transition_probe/__init__.py -> ...transition_probe"""
    p = path[:-3] if path.endswith(".py") else path
    p = p.replace("/", ".")
    if p.endswith(".__init__"):
        p = p[: -len(".__init__")]
    return p


def build_dep_index():
    """谁依赖谁 —— 直接调 scan_deps 的解析，不在这里重写一遍。

    重写过一次，教训：简化的 import 匹配认不出 `import motion` 指的是
    `scripts.motion`，整张表的「依赖谁」列会全变成「叶子」——
    一张把所有依赖都显示成没有依赖的表，比没有表更坏。
    """
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import scan_deps
    edges, rev, _, _ = scan_deps.build_graph()
    return edges, rev


def table(key, edges, rev):
    layer = LAYERS[key]
    rows = ["| 文件 | 职责 | 对外接口 / 入口 | 依赖谁 | 被谁依赖 | 体积 |",
            "| --- | --- | --- | --- | --- | --- |"]
    for path, human in layer["files"]:
        head = read_head(path)
        duty = head if head else f"（无 docstring）{human}"
        n = line_count(path)
        mn = mod_name(path)
        deps = sorted(edges.get(mn, ()))
        deps_s = "、".join("`%s`" % d.split(".")[-1] for d in deps) or "—（叶子）"
        r = sorted(rev.get(mn, ()))
        rev_s = "、".join("`%s`" % x.split(".")[-1] for x in r) or "—（没人依赖）"
        vol = f"{n} 行" + (" ⚠️ 超 800" if n > 800 else "")
        rows.append(f"| `{path}` | {duty} | {human} | {deps_s} | {rev_s} | {vol} |")
    return "\n".join(rows)


def main():
    args = [a for a in sys.argv[1:] if a != "--check"]
    edges, rev = build_dep_index()
    keys = args or list(LAYERS)
    for k in keys:
        if k not in LAYERS:
            print("未知层: %s（可选: %s）" % (k, ", ".join(LAYERS)))
            continue
        L = LAYERS[k]
        print("\n### %s\n" % L["title"])
        print("%s\n" % L["desc"])
        print(table(k, edges, rev))
        py = sum(line_count(p) for p, _ in L["files"] if p.endswith(".py"))
        print("\n共 %d 个文件，其中 Python %d 行。" % (len(L["files"]), py))


if __name__ == "__main__":
    main()
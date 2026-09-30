# 会话交接 · 2026-09-30

> **这份文档为"上一次会话卡住"而写。** 目的只有一个：让下一个接手的人
> （人或 agent）在 **5 分钟内**知道 —— 做到哪了、东西在哪、怎么接着做。
>
> 项目的**长期**交接文档是 [`HANDOVER.md`](HANDOVER.md)（架构、目录、误判史）。
> 这一份只管**这一次会话**。

---

## 0. 一句话状态

**三件事全部做完、测试全绿、已推送远端。本次会话已收尾，无遗留未完成动作。**

| # | 任务 | 状态 |
| --- | --- | --- |
| B-① | `<p:timing>` × 切换 的配合实测 | ✅ 完成 · `92a8ed2` · **在远端** |
| B-② | 48 项之间的取舍 | ✅ 完成 · `8eca1c2` · **在远端** |
| B-③ | 大型文档整理 / 拆分 | ✅ 完成 · `b03d7e7`（远端 tip `eb20f09`）· **已推送** |

> **"卡住"的真相**：不是出错。上一次会话的最后一步是**逐个跑 8 个测试套件**，
> 这需要几分钟，输出被缓冲住了，看起来像死掉。**把它杀掉是安全的** ——
> 工作区没有任何半写状态。
>
> 我这次重跑了一遍：**8/8 全绿**（§3），提交了 B-③（§1），并推送成功（§1b）。

---

## 1. 推送已完成（留档）

```bash
# 走 API 推送（这台机器 github.com:443 被代理挡，git push 会挂 —— 见 §6）
TOKEN=$(printf 'protocol=https\nhost=github.com\n\n' | git credential fill \
        2>/dev/null | sed -n 's/^password=//p')
GH_TOKEN="$TOKEN" REPO_DIR="D:/workbuddy/idea/ppt-motion-skill" \
  "$PY" scripts/push_via_api.py
# → 远端 5dcff09 -> 本地 b03d7e7，共 1 个提交；远端 main 已更新到 eb20f09
```

**用 `tree` 哈希验证内容一致**（⚠️ 不能看 commit sha —— API 造的 sha 必然
与本地不同）：

```bash
git rev-parse HEAD^{tree}
# → ebe980a3e1690f857202762a3bb52706e74f2c4b
# 远端同 sha（经 API 查 /git/commits/<tip> 得到）→ 内容逐字节相同 ✓
```

> **撤销提交**（如果你还想复核拆分）：`git reset --soft HEAD~1` ——
> 改动回到工作区，一个字节都不会丢。**但注意远端已经推上去了。**

---

## 1b. 这次提交里有什么（规模）

`b03d7e7`：**31 个文件**，全在 B-③ 名下。

```
git show --stat b03d7e7 | tail -3
# → 31 files changed
#   reference/com-pitfalls.md  −2040 行（退化成目录页）
#   reference/symptoms.md      −1403 行（退化成目录页）
#   + 16 份拆分文件、映射表、生成器、测试、交接文档
```

---

## 2. 这次会话到底做了什么

### B-①　切换 × 页内动画：推翻了文档里三处断言

**问题**：同一页既有 `<p:transition>`（切换）又有 `<p:timing>`（页内动画），
两者什么关系？文档三处写着"正交、互不占用"。

**实测结论**（9 份探测 deck，逐帧量）：

- **结构上平行** —— 两块都活着、顺序也对（9/9）。这部分旧文档没错。
- **⚠️ 时间上串行** —— 页内动画**要排队等切换演完**才开始。
  **"正交"这个词是错的**，旧文档三处需要按新结论读。

**延迟规律**（本次最值钱的数字）：

| 同页的切换 | 动画延迟 |
| --- | --- |
| 无 | +67 ms |
| `fade` 300 ms | +433 ms |
| `push` / `wipe` 800 ms | +933 ms |
| `fade` 1500 ms | +1633 ms |

→ **延迟 ≈ 切换时长 + 133 ms**。动画**自身长度不变**（换什么切换都是 467 ms）。

**顺带发现的坑**：通用"帧间能量"指标对**淡入完全失明** ——
中灰圆盘淡入中灰网格，整帧均值变化 < 2/255，指标报"零个活跃帧"，
看起来像"动画没生效"。

> ⚠️ **这是同类坑第三次**：§47 旋转躲过位移/漂移/径向三类指标、
> §48 整页平移躲过变化重心、§51 淡入躲过能量阈值。
> **规律：每加一种"看的方式"，都要问它看不见什么。**

**踩的坑（值得记住）**：第一版对照实验里，组合行用 `<p:push/>`、
单跑行用 `<p:fade/>` —— **一次动两个变量**，低能量的 fade 被读成
"切换消失了"。**对照实验必须与组合行用同一个切换元素。**

**产出**：`reference/transition-model.md` 新增 §五、
`reference/pitfall-transition.md` §51、`facts/transitions.json` +3 条规则、
测试第 14 组。

### B-②　48 项之间的取舍：48 个名字 ≠ 48 种选择

**问题**：48 个切换效果，选型时到底怎么挑？

**核心发现**：48 个名字里**真正独立的维度只有三个** ——

1. **方向轴**（横 / 竖 / 对称）—— 族内第一道筛
2. **`dir` 能不能拧** —— 22 个方向揭示成员里**只有 8 个**会镜像
3. **帧跨度**（运动窗口长度）—— 唯一能当"时长预算"用的量

其余差别是**观感**。**48 个名字里 40 多个是"同一件事的不同皮肤"。**

**两条硬约束**：

- `box`（XML 是 `<p:zoom>`）与 `comb` 加 `dir` 会让 **PowerPoint 打不开整份文件**。
- `clock` 的跨度记作 4，但那是**指标看不见旋转**的伪值（§47），**不可拿它比大小**。

**产出**：`reference/transition-choice.md` 新增 §二·补（补-1 ~ 补-6）、
测试第 15 组。

### B-③　大型文档拆分（已提交 `b03d7e7` · 已推送）

**触发**：用户观察到"已经堆积到 §50 了"。当时 `com-pitfalls.md` 2027 行 / 51 节，
`symptoms.md` 1394 行。

**拆法**：

| 原文件 | 拆成 | 拆的依据 | 引用怎么活下来 |
| --- | --- | --- | --- |
| `com-pitfalls.md` | **10 份** `pitfall-*.md` | 按**主题** | **编号不变** —— `§44` 永远指"切换的形态"；编号→文件表在 `pitfall-map.md`（**脚本扫描生成，不是手抄**） |
| `symptoms.md` | **6 份** `symptom-*.md` | 按**现象** | **栏名不变** —— 原目录页保留同名分流 |

**两条设计取舍（这是本次最值得记的判断）**：

1. **原文件名必须保留**，退化成目录页，**不是删除**。全仓库 **91 处**引用
   写着 `com-pitfalls.md`。删掉 = 91 个死链。**改引用比留入口贵得多，也更容易漏。**
2. **⚠️ 编号只增不减、只挪不改。** 一个编号永远指同一件事，哪怕它换了文件。
   编号一旦被复用，历史引用会**静默**指向另一个坑 —— **比 404 更糟**，
   因为它看起来是好的。

**顺带修掉的一个真 bug**（`scripts/verify_docs.py`）：
孤儿检查拿**链接原文**直接和仓库根路径比，而链接是**相对**的 ——
`reference/com-pitfalls.md` 里写的是 `pitfall-com.md`，字符串永远不等。
之前没暴露是因为所有文档都从 INDEX 直连；一旦出现"被同目录兄弟文件引用"
（拆分就是这么干的），真正可达的文件会被误报成孤儿。
**修法：按链接所在目录解析。**

**产出**：20 个新文件 + 测试第 16/17 组 + `HANDOVER.md` §3.1 记下这条规模规则。

---

## 3. 验证状态（我刚重跑过，全绿）

```
test_transition_table.py    切换表测试 PASSED        ← 17 组
test_install_manifest.py    install manifest ok (11 项)
test_camera.py              camera test passed
test_morph.py               morph test passed
test_singletons.py          15 passed, 0 failed
test_dual_photo.py          23 passed, 0 failed
test_fill_window.py         28 passed, 0 failed
smoke.py                    SMOKE PASSED

scripts/verify_docs.py      OK 引用可解析、无孤儿、入口在预算内
                            INDEX.md 5722 字符（预算 6000）
scripts/build_pitfall_map.py --check
                            OK 映射表与 10 个文件一致（51 条编号）
```

**拆分无损的证明**（逐字节，不是抽查）：

- 51 个踩坑节：与拆分前 `git show HEAD:reference/com-pitfalls.md` 逐节比对，
  **编号集合相同、内容不一致的节 = 0**
- 10 个症状栏：同样的方法，**丢失 0 栏、多出 0 栏、内容不一致 = 0**

> 复现命令见 §6。**改完文档拆分后应该重跑这两条**，这是唯一能证明"没丢东西"的检查。

---

## 4. B-③ 的提交信息（已落库，留档）

```
refactor(docs): 拆分两份巨型文档 —— 编号是接口，不是正文

com-pitfalls.md 2027 行 / 51 节、symptoms.md 1394 行，都到了"查它要翻屏"
的规模。查这些文档的时刻通常是你已经出错的时候，那时候最不该付出的
成本是找。按主题 / 现象各拆成多份。

两条设计取舍：
  1. 原文件名保留成目录页，不删 —— 全仓库 91 处引用写着 com-pitfalls.md，
     删掉就是 91 个死链。改引用比留入口贵，也更容易漏。
  2. 编号只增不减、只挪不改 —— §44 永远指"切换的形态"，哪怕它换了文件。
     编号一旦复用，历史引用会静默指向另一个坑，比 404 更糟。

拆分无损（逐字节验证，非抽查）：
  - 51 个踩坑节：编号集合相同，内容不一致 = 0
  - 10 个症状栏：丢失 0 / 多出 0 / 内容不一致 = 0

顺带修掉 verify_docs.py 的孤儿误报：它拿链接原文和仓库根路径直接比，
而链接是相对的。之前没暴露是因为所有文档都从 INDEX 直连；一旦出现
"被同目录兄弟文件引用"就会把可达文件报成孤儿。改为按链接所在目录解析。

新增：
  - scripts/build_pitfall_map.py  编号→文件映射表生成器（表不能手抄，会烂）
  - tests 第 16/17 组：编号完整、唯一、映射表与真实归属一致、目录页链全
  - tools/split_*.py 一次性拆分脚本（DEV_ONLY，不随包分发）
  - HANDOVER.md §3.1 记下"规模分层"这条规则
```

---

## 5. 剩下要做的

### ✅ 本次会话的动作已全部完成

提交（`b03d7e7`）+ 推送（远端 tip `eb20f09`，tree 哈希一致）。

### 下次接手时可做的（都不急）

#### ⛔ 用户明确搁置、**不要碰**

- `building/` / `template/` 的 **5 个悬空证据项**
  （`verify_recipes.py` 报 exit=1）。**这是已知项，用户决定先放着。**
- **封面占位符**。同样已决定搁置。

#### 顺路可做

- `INDEX.md` 现在 **5722 / 6000 字符** —— 上次差点超预算
  （我一度写重复表冲到 6062 被拦下）。**下次再加内容前先想怎么腾地方。**
- `tools/` 是我新建的**一次性脚本**落脚处，已归入 `test_install_manifest.py`
  的 `DEV_ONLY`（不随技能包分发）。**这条界线要保持**：
  `scripts/` = 运行时工具（要分发），`tools/` = 开发期一次性（不分发）。

---

## 6. 怎么复现 / 怎么接着做

**环境**（本机固定路径）：

```bash
PY="C:/Users/Administrator/.workbuddy/binaries/python/envs/default/Scripts/python.exe"
cd D:/workbuddy/idea/ppt-motion-skill
```

**跑测试**（逐个跑，别用 `unittest discover` —— 它会因 test_dual_photo /
test_fill_window / test_singletons 报 3 个 loader error，**属正常**）：

```bash
for t in tests/test_*.py tests/smoke.py; do
  echo "[$t] $("$PY" "$t" 2>&1 | tail -1)"
done
```

**重证拆分无损**（改过文档拆分后必跑）：

```bash
"$PY" - <<'EOF'
import subprocess, io, re, os
head = subprocess.check_output(['git','rev-parse','HEAD']).decode().strip()
old = subprocess.check_output(
    ['git','show','%s:reference/com-pitfalls.md' % head]).decode('utf-8')

def secs(t):
    ms = list(re.finditer(r'^## (?!#)(\d+)\. (.+)$', t, re.M))
    o = {}
    for i, m in enumerate(ms):
        end = ms[i+1].start() if i+1 < len(ms) else len(t)
        o[int(m.group(1))] = t[m.start():end].rstrip()
    return o

a = secs(old); b = {}
for f in os.listdir('reference'):
    if f.startswith('pitfall-'):
        b.update(secs(io.open('reference/'+f, encoding='utf-8').read()))
print('旧 %d / 新 %d，集合相同 %s，不一致 %s'
      % (len(a), len(b), set(a) == set(b),
         [n for n in a if a[n] != b.get(n)] or '无'))
EOF
```

**文档结构校验**：

```bash
"$PY" scripts/verify_docs.py
"$PY" scripts/build_pitfall_map.py --check   # 改了拆分文件后
```

**推送**（⚠️ 这台机器 `github.com:443` 被代理挡，`git push` 会挂）：

```bash
TOKEN=$(printf 'protocol=https\nhost=github.com\n\n' | git credential fill \
        2>/dev/null | sed -n 's/^password=//p')
GH_TOKEN="$TOKEN" REPO_DIR="D:/workbuddy/idea/ppt-motion-skill" \
  "$PY" scripts/push_via_api.py
```

> **判断推送成败看 `tree` 哈希，不要看 commit sha** —— API 造的 commit
> sha 必然与本地不同（详见 `pitfall-tooling.md` §45）。
>
> 我这次就是这么推的 B-①②：本地 `8eca1c2` → 远端 `5dcff09`，
> **tree 哈希一致**（`da93b95…`）即内容逐字节相同。
>
> ⚠️ 因此 `git rev-list --count origin/main..HEAD` 会常年显示
> `ahead=2` 之类的**假数字** —— 那两个提交其实已在远端。别被它骗了。

---

## 7. 五条最值得带走的教训

1. **对照实验一次只动一个变量。** 我第一版让组合行和单跑行用了不同切换元素，
   结论直接反了。
2. **每加一种"看的方式"，都要问它看不见什么。** 旋转、整页平移、淡入
   各自躲过了不同的指标 —— 同类坑已发生三次。
3. **编号是接口。** 拆分文档时先问"有多少外部引用指着它"，
   再决定保留什么。91 处引用让"删掉旧文件名"这个直觉选项直接出局。
4. **"卡住"往往不是出错，是等待。** 这次"卡住"= 8 个测试套件跑几分钟
   输出被缓冲。**杀进程前先看工作区有没有半写状态** —— 没有就可以放心。
5. **"我写的不生效" ≠ "这功能不支持"。** 先让 PowerPoint 自己写一遍，
   比读文档或推理都快。

---

*本文档由会话收尾时生成。项目长期交接见 [`HANDOVER.md`](HANDOVER.md)。*

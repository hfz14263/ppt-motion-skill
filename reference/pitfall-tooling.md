# 踩坑 · 工具链

> 工具链 —— 推送通道、导出、视频、运行方式
>
> **编号沿用拆分前的 `§n`**（`§44` 仍是"切换的形态"，哪怕它换了文件）。
> 完整编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 按现象查见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

## 32. `CreateVideo` 报"完成"却拿不到文件 —— 三个参数全错

给真实文档导 MP4（媒体层），第一遍调用：

```powershell
$pres = $ppt.Presentations.Open($deck, -1, 0, 0)   # WithWindow = 0
$pres.CreateVideo($mp4, $true, 6)                   # 只给 3 个参数
while ($pres.CreateVideoStatus -eq 1) { ... }       # 等状态离开 InProgress
```

结果：**轮询立刻退出，日志写 `status=2 exists=True`，几秒后文件不存在**。
我当时把 `2` 读成"完成"——实际 `PpMediaTaskStatus` 里 **2 = Failed，3 = Done**。
`exists=True` 是编码器刚落了个空壳文件，Quit 之后被清掉了。三处错叠加：

| 错 | 后果 |
| --- | --- |
| `WithWindow = 0` 打开 | 编码器起不来（probe 脚本用的是 `WithWindow = 1`） |
| 没重定向 `TEMP/TMP` | 沙箱 TEMP 会让编码器找不到可写的临时目录 |
| `$pres.CreateVideo($f,$true,6)` 直接调 | 部分构建上late-binding 丢参，要 `.Invoke(@($f,$true,6,720,30,85))` 六参全给 |

按 `scripts/probe_createvideo.ps1` 的姿势重试：**开窗打开 + TEMP 重定向 +
六参 Invoke + 轮询到 `status -eq 3 -or -eq 4`** —— 同一台机器 15 秒出片
（720p/30fps/q85，11.8 MB / 78 秒成片）。

### 规则

1. `CreateVideoStatus`：**0/1 = 没开始/进行中，2 = 失败，3/4 = 完成**。
   按 `2` 判完成会把失败当成品。
2. 先跑 `probe_createvideo.ps1` 测**当前机器**再写业务脚本 —— 该能力是
   构建/驱动相关的，别信任何一台机器的经验（包括自己昨天那台）。
3. `player.py` 依赖 `win32com`，装不上 pywin32 的环境里 MP4 就是唯一的
   播放证据 —— 这条链路值得保通。
4. **输出路径别超过 255 字符。** 同样的脚本，写到深层工作目录（长中文路径）
   时 `CreateVideo` 静默不产文件，换到 `C:\t1w\mine.mp4` 这种短路径立刻成功。
   排查"参数都对却没文件"时，**先把输出路径改短再复测**，别急着怀疑 Office 构建。

---
## 45. 推送：github.com 被挡时，改走 api.github.com

### 45.1 症状

`git push` 报：

```
fatal: unable to access 'https://github.com/...': CONNECT tunnel failed, response 502
```

但同一时刻 `curl -x <proxy> https://api.github.com` **通（200）**。
逐个域名测下来：

| 域名 | 代理结果 |
|---|---|
| `api.github.com` | **200** |
| `codeload.github.com` | 301 |
| `raw.githubusercontent.com` | 301 |
| `github.com` | **000（被挡）** |
| `gist.github.com` | 000 |

即：本机代理放行了 API 与静态资源，唯独不放行 `github.com:443`。
而 `git push` 走的是后者。**换 PAT、重试、加 `-c http.proxy=` 都没用** —— 不是凭证问题，
是通道问题。

### 45.2 走法：Git Data API 三代（blob → tree → commit → 更新 ref）

`api.github.com` 通，就能手工把 commit 造到远端：

```
POST /repos/{o}/{r}/git/blobs     {content: base64, encoding: "base64"}
POST /repos/{o}/{r}/git/trees     {tree: [{path, mode, type:"blob", sha}...]}
POST /repos/{o}/{r}/git/commits   {message, tree, parents, author}
PATCH /repos/{o}/{r}/git/refs/heads/main  {sha, force: false}
```

`GET /repos/{o}/{r}/commits/main` 拿远端当前 tree 做核对（见 45.4）。

### 45.3 ⚠️ 最大的坑：必须上传 git 对象字节，不能读工作区文件

本仓库 `core.autocrlf=true`：

```
blob 字节: 63742   有 CRLF: 0     ← 对象库里存的是 LF
工作区字节: 64981   有 CRLF: 1239  ← 工作区是 CRLF
```

直接 `open(path,'rb').read()` 上传，GitHub 按 CRLF 内容算出**另一个 sha**
（本地 `0f3fb792…`，GitHub 返回 `7c32fd9a…`）。随后建 tree 时：

```
422 tree.sha 0f3fb792... is not a valid blob
```

**正确做法**：先 `git cat-file blob <sha>` 取字节再 base64 —— 上传后校验
`res["sha"] == bsha`，不等就立刻炸，别让它攒到建 tree 时才报错。

### 45.4 正确性判据：比 tree 哈希，不比 commit sha

API 造的 commit，**sha 必然与本地不同**（committer/时间戳不同），
所以 `git log origin/main..HEAD` 会一直显示"ahead 2" —— 这不代表没推上去。

真正的判据是 **tree 哈希**：

```
本地  git rev-parse HEAD^{tree}   → 889dbef934dcd811a3fa83c6c8a005e7712f8012
远端  commits/main.commit.tree.sha → 889dbef934dcd811a3fa83c6c8a005e7712f8012
```

相等 = 内容逐字节一致。再补一条 `git diff HEAD origin/main`（应为空），
然后 `git reset --hard origin/main` 把本地对齐，收工。

### 45.5 别用 GET /git/blobs/<sha> 判断"要不要上传"

该接口只认仓库**已有对象**，对新增但内容相同的 blob 也返回成功或 404，
判断不可靠。改用「父 commit 的 tree 里是否已有 `path|sha`」做差集。

---

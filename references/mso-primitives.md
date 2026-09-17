# 动画效果索引

`motion.py catalog` 是权威来源；本文件是离线速查。

所有 `presetID / presetClass / presetSubtype` 都是**从真实 PowerPoint 输出提取**的
（用 COM 逐个 `AddEffect` 后回读 XML），不是按枚举值猜的。

## 重要：enum ≠ presetID

多数情况下两者相同，但**有例外**，所以注入必须用 presetID：

| 别名 | MsoAnimEffect 枚举 | 真实 presetID |
| --- | --- | --- |
| `changeFillColor` | 54 | **1** |
| `pathCircle` | 86 | **1** |
| `changeFont` | 55 | **2** |
| `pathRightTriangle` | 87 | **2** |
| `changeFontColor` | 56 | **3** |
| `pathDiamond` | 88 | **3** |
| `changeFontSize` | 57 | **4** |
| `pathHexagon` | 89 | **4** |
| `changeFontStyle` | 58 | **5** |
| `path5PointStar` | 90 | **5** |
| `growShrink` | 59 | **6** |
| `pathCrescentMoon` | 91 | **6** |
| `changeLineColor` | 60 | **7** |
| `pathSquare` | 92 | **7** |
| `spin` | 61 | **8** |
| `pathTrapezoid` | 93 | **8** |
| `transparency` | 62 | **9** |
| `pathHeart` | 94 | **9** |
| `boldFlash` | 63 | **10** |
| `pathOctagon` | 95 | **10** |
| `path6PointStar` | 96 | **11** |
| `pathFootball` | 97 | **12** |
| `pathEqualTriangle` | 98 | **13** |
| `blast` | 64 | **14** |
| `pathParallelogram` | 99 | **14** |
| `boldReveal` | 65 | **15** |
| `pathPentagon` | 100 | **15** |
| `brushOnColor` | 66 | **16** |
| `path4PointStar` | 101 | **16** |
| `path8PointStar` | 102 | **17** |
| `brushOnUnderline` | 67 | **18** |
| `pathTeardrop` | 103 | **18** |
| `colorBlend` | 68 | **19** |
| `pathPointyStar` | 104 | **19** |
| `colorWave` | 69 | **20** |
| `pathCurvedSquare` | 105 | **20** |
| `complementaryColor` | 70 | **21** |
| `pathCurvedX` | 106 | **21** |
| `pathVerticalFigure8` | 107 | **22** |
| `contrastColor` | 72 | **23** |
| `pathCurvyStar` | 108 | **23** |
| `darken` | 73 | **24** |
| `pathLoopdeLoop` | 109 | **24** |
| `desaturate` | 74 | **25** |
| `pathBuzzsaw` | 110 | **25** |
| `flashBulb` | 75 | **26** |
| `pathHorizontalFigure8` | 111 | **26** |
| `flicker` | 76 | **27** |
| `pathPeanut` | 112 | **27** |
| `growWithColor` | 77 | **28** |
| `pathFigure8Four` | 113 | **28** |
| `pathNeutron` | 114 | **29** |
| `lighten` | 78 | **30** |
| `pathSwoosh` | 115 | **30** |
| `pathBean` | 116 | **31** |
| `teeter` | 80 | **32** |
| `pathPlus` | 117 | **32** |
| `verticalGrow` | 81 | **33** |
| `pathInvertedTriangle` | 118 | **33** |
| `lightSpeed` | 32 | **34** |
| `wave` | 82 | **34** |
| `pathInvertedSquare` | 119 | **34** |
| `pinwheel` | 33 | **35** |
| `pathLeft` | 120 | **35** |
| `shimmer` | 52 | **36** |
| `pathTurnRight` | 121 | **36** |
| `riseUp` | 34 | **37** |
| `pathArcDown` | 122 | **37** |
| `swish` | 35 | **38** |
| `pathZigzag` | 123 | **38** |
| `pathSCurve2` | 124 | **39** |
| `unfold` | 37 | **40** |
| `pathSineWave` | 125 | **40** |
| `whip` | 38 | **41** |
| `pathBounceLeft` | 126 | **41** |
| `ascend` | 39 | **42** |
| `pathDown` | 127 | **42** |
| `centerRevolve` | 40 | **43** |
| `pathTurnUp` | 128 | **43** |
| `pathArcUp` | 129 | **44** |
| `pathHeartbeat` | 130 | **45** |
| `pathSpiralRight` | 131 | **46** |
| `descend` | 42 | **47** |
| `pathWave` | 132 | **47** |
| `sling` | 43 | **48** |
| `pathCurvyLeft` | 133 | **48** |
| `spinner` | 44 | **49** |
| `pathDiagonalDownRight` | 134 | **49** |
| `pathTurnDown` | 135 | **50** |
| `zip` | 46 | **51** |
| `pathArcLeft` | 136 | **51** |
| `arcUp` | 47 | **52** |
| `pathFunnel` | 137 | **52** |
| `fadedZoom` | 48 | **53** |
| `pathSpring` | 138 | **53** |
| `glide` | 49 | **54** |
| `pathBounceRight` | 139 | **54** |
| `expand` | 50 | **55** |
| `pathSpiralLeft` | 140 | **55** |
| `flip` | 51 | **56** |
| `pathDiagonalUpRight` | 141 | **56** |
| `pathTurnUpRight` | 142 | **57** |
| `fold` | 53 | **58** |
| `pathArcRight` | 143 | **58** |
| `pathSCurve1` | 144 | **59** |
| `pathDecayingWave` | 145 | **60** |
| `pathCurvyRight` | 146 | **61** |
| `pathStairsDown` | 147 | **62** |
| `pathRight` | 149 | **63** |
| `pathUp` | 148 | **64** |

## 已收录（137 个，可直接用）

标记：`P`=带动作路径 `F`=带 animEffect 滤镜 `V`=带 style.visibility 设置

### 入场（45）

| 别名 | presetID | subtype | 特性 | MsoAnimEffect |
| --- | --- | --- | --- | --- |
| `appear` | 1 | 0 | `··V` | msoAnimEffectAppear |
| `fly` | 2 | 4 | `··V` | msoAnimEffectFly |
| `blinds` | 3 | 10 | `·FV` | msoAnimEffectBlinds |
| `box` | 4 | 16 | `·FV` | msoAnimEffectBox |
| `checkerboard` | 5 | 10 | `·FV` | msoAnimEffectCheckerboard |
| `circle` | 6 | 16 | `·FV` | msoAnimEffectCircle |
| `crawl` | 7 | 4 | `··V` | msoAnimEffectCrawl |
| `diamond` | 8 | 16 | `·FV` | msoAnimEffectDiamond |
| `dissolve` | 9 | 0 | `·FV` | msoAnimEffectDissolve |
| `fade` | 10 | 0 | `·FV` | msoAnimEffectFade |
| `flash` | 11 | 0 | `··V` | msoAnimEffectFlashOnce |
| `peek` | 12 | 4 | `·FV` | msoAnimEffectPeek |
| `plus` | 13 | 16 | `·FV` | msoAnimEffectPlus |
| `randomBars` | 14 | 10 | `·FV` | msoAnimEffectRandomBars |
| `spiral` | 15 | 0 | `··V` | msoAnimEffectSpiral |
| `split` | 16 | 21 | `·FV` | msoAnimEffectSplit |
| `stretch` | 17 | 10 | `··V` | msoAnimEffectStretch |
| `strips` | 18 | 12 | `·FV` | msoAnimEffectStrips |
| `swivel` | 19 | 10 | `··V` | msoAnimEffectSwivel |
| `wedge` | 20 | 0 | `·FV` | msoAnimEffectWedge |
| `wheel` | 21 | 1 | `·FV` | msoAnimEffectWheel |
| `wipe` | 22 | 4 | `·FV` | msoAnimEffectWipe |
| `zoom` | 23 | 16 | `··V` | msoAnimEffectZoom |
| `bounce` | 26 | 0 | `·FV` | msoAnimEffectBounce |
| `credits` | 28 | 0 | `··V` | msoAnimEffectCredits |
| `float` | 30 | 0 | `·FV` | msoAnimEffectFloat |
| `growTurn` | 31 | 0 | `·FV` | msoAnimEffectGrowAndTurn |
| `lightSpeed` | 34 | 0 | `··V` | msoAnimEffectLightSpeed |
| `pinwheel` | 35 | 0 | `·FV` | msoAnimEffectPinwheel |
| `riseUp` | 37 | 0 | `·FV` | msoAnimEffectRiseUp |
| `swish` | 38 | 0 | `··V` | msoAnimEffectSwish |
| `unfold` | 40 | 0 | `·FV` | msoAnimEffectUnfold |
| `whip` | 41 | 0 | `·FV` | msoAnimEffectWhip |
| `ascend` | 42 | 0 | `·FV` | msoAnimEffectAscend |
| `centerRevolve` | 43 | 0 | `·FV` | msoAnimEffectCenterRevolve |
| `descend` | 47 | 0 | `·FV` | msoAnimEffectDescend |
| `sling` | 48 | 0 | `·FV` | msoAnimEffectSling |
| `spinner` | 49 | 0 | `·FV` | msoAnimEffectSpinner |
| `zip` | 51 | 0 | `·FV` | msoAnimEffectZip |
| `arcUp` | 52 | 0 | `PFV` | msoAnimEffectArcUp |
| `fadedZoom` | 53 | 16 | `·FV` | msoAnimEffectFadedZoom |
| `glide` | 54 | 0 | `·FV` | msoAnimEffectGlide |
| `expand` | 55 | 0 | `·FV` | msoAnimEffectExpand |
| `flip` | 56 | 0 | `··V` | msoAnimEffectFlip |
| `fold` | 58 | 0 | `·FV` | msoAnimEffectFold |

### 强调（28）

| 别名 | presetID | subtype | 特性 | MsoAnimEffect |
| --- | --- | --- | --- | --- |
| `changeFillColor` | 1 | 2 | `···` | msoAnimEffectChangeFillColor |
| `changeFont` | 2 | 0 | `···` | msoAnimEffectChangeFont |
| `changeFontColor` | 3 | 2 | `···` | msoAnimEffectChangeFontColor |
| `changeFontSize` | 4 | 2 | `···` | msoAnimEffectChangeFontSize |
| `changeFontStyle` | 5 | 1 | `···` | msoAnimEffectChangeFontStyle |
| `growShrink` | 6 | 0 | `···` | msoAnimEffectGrowShrink |
| `changeLineColor` | 7 | 2 | `···` | msoAnimEffectChangeLineColor |
| `spin` | 8 | 0 | `···` | msoAnimEffectSpin |
| `transparency` | 9 | 0 | `·F·` | msoAnimEffectTransparency |
| `boldFlash` | 10 | 0 | `···` | msoAnimEffectBoldFlash |
| `blast` | 14 | 0 | `···` | msoAnimEffectBlast |
| `boldReveal` | 15 | 0 | `···` | msoAnimEffectBoldReveal |
| `brushOnColor` | 16 | 0 | `···` | msoAnimEffectBrushOnColor |
| `brushOnUnderline` | 18 | 0 | `···` | msoAnimEffectBrushOnUnderline |
| `colorBlend` | 19 | 0 | `···` | msoAnimEffectColorBlend |
| `colorWave` | 20 | 0 | `···` | msoAnimEffectColorWave |
| `complementaryColor` | 21 | 0 | `···` | msoAnimEffectComplementaryColor |
| `contrastColor` | 23 | 0 | `···` | msoAnimEffectContrastingColor |
| `darken` | 24 | 0 | `···` | msoAnimEffectDarken |
| `desaturate` | 25 | 0 | `···` | msoAnimEffectDesaturate |
| `flashBulb` | 26 | 0 | `·F·` | msoAnimEffectFlashBulb |
| `flicker` | 27 | 0 | `···` | msoAnimEffectFlicker |
| `growWithColor` | 28 | 0 | `···` | msoAnimEffectGrowWithColor |
| `lighten` | 30 | 0 | `···` | msoAnimEffectLighten |
| `teeter` | 32 | 0 | `···` | msoAnimEffectTeeter |
| `verticalGrow` | 33 | 0 | `···` | msoAnimEffectVerticalGrow |
| `wave` | 34 | 0 | `P··` | msoAnimEffectWave |
| `shimmer` | 36 | 0 | `···` | msoAnimEffectShimmer |

### 动作路径（64）

| 别名 | presetID | subtype | 特性 | MsoAnimEffect |
| --- | --- | --- | --- | --- |
| `pathCircle` | 1 | 0 | `P··` | msoAnimEffectPathCircle |
| `pathRightTriangle` | 2 | 0 | `P··` | msoAnimEffectPathRightTriangle |
| `pathDiamond` | 3 | 0 | `P··` | msoAnimEffectPathDiamond |
| `pathHexagon` | 4 | 0 | `P··` | msoAnimEffectPathHexagon |
| `path5PointStar` | 5 | 0 | `P··` | msoAnimEffectPath5PointStar |
| `pathCrescentMoon` | 6 | 0 | `P··` | msoAnimEffectPathCrescentMoon |
| `pathSquare` | 7 | 0 | `P··` | msoAnimEffectPathSquare |
| `pathTrapezoid` | 8 | 0 | `P··` | msoAnimEffectPathTrapezoid |
| `pathHeart` | 9 | 0 | `P··` | msoAnimEffectPathHeart |
| `pathOctagon` | 10 | 0 | `P··` | msoAnimEffectPathOctagon |
| `path6PointStar` | 11 | 0 | `P··` | msoAnimEffectPath6PointStar |
| `pathFootball` | 12 | 0 | `P··` | msoAnimEffectPathFootball |
| `pathEqualTriangle` | 13 | 0 | `P··` | msoAnimEffectPathEqualTriangle |
| `pathParallelogram` | 14 | 0 | `P··` | msoAnimEffectPathParallelogram |
| `pathPentagon` | 15 | 0 | `P··` | msoAnimEffectPathPentagon |
| `path4PointStar` | 16 | 0 | `P··` | msoAnimEffectPath4PointStar |
| `path8PointStar` | 17 | 0 | `P··` | msoAnimEffectPath8PointStar |
| `pathTeardrop` | 18 | 0 | `P··` | msoAnimEffectPathTeardrop |
| `pathPointyStar` | 19 | 0 | `P··` | msoAnimEffectPathPointyStar |
| `pathCurvedSquare` | 20 | 0 | `P··` | msoAnimEffectPathCurvedSquare |
| `pathCurvedX` | 21 | 0 | `P··` | msoAnimEffectPathCurvedX |
| `pathVerticalFigure8` | 22 | 0 | `P··` | msoAnimEffectPathVerticalFigure8 |
| `pathCurvyStar` | 23 | 0 | `P··` | msoAnimEffectPathCurvyStar |
| `pathLoopdeLoop` | 24 | 0 | `P··` | msoAnimEffectPathLoopdeLoop |
| `pathBuzzsaw` | 25 | 0 | `P··` | msoAnimEffectPathBuzzsaw |
| `pathHorizontalFigure8` | 26 | 0 | `P··` | msoAnimEffectPathHorizontalFigure8 |
| `pathPeanut` | 27 | 0 | `P··` | msoAnimEffectPathPeanut |
| `pathFigure8Four` | 28 | 0 | `P··` | msoAnimEffectPathFigure8Four |
| `pathNeutron` | 29 | 0 | `P··` | msoAnimEffectPathNeutron |
| `pathSwoosh` | 30 | 0 | `P··` | msoAnimEffectPathSwoosh |
| `pathBean` | 31 | 0 | `P··` | msoAnimEffectPathBean |
| `pathPlus` | 32 | 0 | `P··` | msoAnimEffectPathPlus |
| `pathInvertedTriangle` | 33 | 0 | `P··` | msoAnimEffectPathInvertedTriangle |
| `pathInvertedSquare` | 34 | 0 | `P··` | msoAnimEffectPathInvertedSquare |
| `pathLeft` | 35 | 0 | `P··` | msoAnimEffectPathLeft |
| `pathTurnRight` | 36 | 0 | `P··` | msoAnimEffectPathTurnRight |
| `pathArcDown` | 37 | 0 | `P··` | msoAnimEffectPathArcDown |
| `pathZigzag` | 38 | 0 | `P··` | msoAnimEffectPathZigzag |
| `pathSCurve2` | 39 | 0 | `P··` | msoAnimEffectPathSCurve2 |
| `pathSineWave` | 40 | 0 | `P··` | msoAnimEffectPathSineWave |
| `pathBounceLeft` | 41 | 0 | `P··` | msoAnimEffectPathBounceLeft |
| `pathDown` | 42 | 0 | `P··` | msoAnimEffectPathDown |
| `pathTurnUp` | 43 | 0 | `P··` | msoAnimEffectPathTurnUp |
| `pathArcUp` | 44 | 0 | `P··` | msoAnimEffectPathArcUp |
| `pathHeartbeat` | 45 | 0 | `P··` | msoAnimEffectPathHeartbeat |
| `pathSpiralRight` | 46 | 0 | `P··` | msoAnimEffectPathSpiralRight |
| `pathWave` | 47 | 0 | `P··` | msoAnimEffectPathWave |
| `pathCurvyLeft` | 48 | 0 | `P··` | msoAnimEffectPathCurvyLeft |
| `pathDiagonalDownRight` | 49 | 0 | `P··` | msoAnimEffectPathDiagonalDownRight |
| `pathTurnDown` | 50 | 0 | `P··` | msoAnimEffectPathTurnDown |
| `pathArcLeft` | 51 | 0 | `P··` | msoAnimEffectPathArcLeft |
| `pathFunnel` | 52 | 0 | `P··` | msoAnimEffectPathFunnel |
| `pathSpring` | 53 | 0 | `P··` | msoAnimEffectPathSpring |
| `pathBounceRight` | 54 | 0 | `P··` | msoAnimEffectPathBounceRight |
| `pathSpiralLeft` | 55 | 0 | `P··` | msoAnimEffectPathSpiralLeft |
| `pathDiagonalUpRight` | 56 | 0 | `P··` | msoAnimEffectPathDiagonalUpRight |
| `pathTurnUpRight` | 57 | 0 | `P··` | msoAnimEffectPathTurnUpRight |
| `pathArcRight` | 58 | 0 | `P··` | msoAnimEffectPathArcRight |
| `pathSCurve1` | 59 | 0 | `P··` | msoAnimEffectPathSCurve1 |
| `pathDecayingWave` | 60 | 0 | `P··` | msoAnimEffectPathDecayingWave |
| `pathCurvyRight` | 61 | 0 | `P··` | msoAnimEffectPathCurvyRight |
| `pathStairsDown` | 62 | 0 | `P··` | msoAnimEffectPathStairsDown |
| `pathRight` | 63 | 0 | `P··` | msoAnimEffectPathRight |
| `pathUp` | 64 | 0 | `P··` | msoAnimEffectPathUp |

## 尚未收录的 MsoAnimEffect 值

这些效果没有从 PowerPoint 提取模板，注入会报 unknown effect alias。要补：把枚举值加进 `scripts/alias_probe.ps1` 的映射表，重跑提取。

| 枚举 | 名称 |
| --- | --- |
| 0 | msoAnimEffectCustom |
| 24 | msoAnimEffectRandomEffects |
| 25 | msoAnimEffectBoomerang |
| 27 | msoAnimEffectColorReveal |
| 29 | msoAnimEffectEaseIn |
| 36 | msoAnimEffectThinLine |
| 41 | msoAnimEffectFadedSwivel |
| 45 | msoAnimEffectStretchy |
| 71 | msoAnimEffectComplementaryColor2 |
| 79 | msoAnimEffectStyleEmphasis |
| 83 | msoAnimEffectMediaPlay |
| 84 | msoAnimEffectMediaPause |
| 85 | msoAnimEffectMediaStop |
| 150 | msoAnimEffectMediaPlayFromBookmark |

## 触发与时序

| spec 写法 | PowerPoint nodeType | 语义 |
| --- | --- | --- |
| `trigger: click` | `clickEffect` | 点击触发 |
| `trigger: with` | `withEffect` | 与上一个同时 |
| `trigger: after` | `afterEffect` | 上一个之后（默认） |

其它时序字段：`delay`（秒，写到外层 stCondLst）、`duration`（秒，写到叶子行为 cTn 的毫秒）、
`repeat`（次数，`repeatCount = 次数×1000`）、`autoReverse`（`autoRev="1"`）、
`smooth`（0–1，`accel`/`decel`）。

## 切换效果

以写出的 `<p:transition>` 元素为准。COM 的 `SlideShowTransition.EntryEffect` 枚举在本机
与元素**不一致**（`0x0A01` 实际产出 `<p:strips/>`），故仅用于复核读数。

| spec 名 | 写出元素 | COM 复核枚举 |
| --- | --- | --- |
| `fade` | `<p:fade/>` | 0x0A01 |
| `fadeblack` | `<p:fade thruBlk="1"/>` | 0x0B01 |
| `push` | `<p:push dir="u"/>` | 0x0901 |
| `pushleft` | `<p:push dir="l"/>` | 0x0901 |
| `wipe` | `<p:wipe dir="l"/>` | 0x0801 |
| `cover` | `<p:cover dir="l"/>` | 0x0C01 |
| `split` | `<p:split orient="horz" dir="out"/>` | 0x0701 |
| `zoom` | `<p:zoom dir="in"/>` | 0x0C01 |
| `dissolve` | `<p:dissolve/>` | 0x0D01 |
| `strips` | `<p:strips/>` | 0x0A01 |
| `pull` | `<p:pull/>` | 0x0801 |
| `randombar` | `<p:randomBar/>` | 0x0901 |

# 踩坑 · OOXML 注入写法

> OOXML 注入写法 —— 元素顺序、属性、命名空间
>
> **编号沿用拆分前的 `§n`**（`§44` 仍是"切换的形态"，哪怕它换了文件）。
> 完整编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 按现象查见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

## 3. lxml 不能直接 parse 带编码声明的 str

**现象**：`ET.fromstring()` 抛 `ValueError: Unicode strings with encoding declaration are not supported`

```python
ET.fromstring(z.read(part).decode('utf-8'))        # ValueError
ET.fromstring(z.read(part))                        # bytes，正常
```
## 8. `p14:dur` 需要声明前缀

**现象**：写 `p14:dur="800"` 后 lxml 报未绑定前缀，整份文件非法

想带毫秒时长要写 `p14:dur="800"`，但 pptd 导出的 slide 根元素**没有** `xmlns:p14`。
必须自己补上，否则是未绑定前缀（非法 XML）：

```xml
<p:sld … xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
```

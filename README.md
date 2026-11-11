# dsh-desktop-toolkit · Windows 桌面操控

给 agent 用的桌面操作工具：**看得见、点得准、有证据**。

```
窗口作用域   锁定一个窗口，落在窗口外的点击一律拒绝
命名锚点     anchor set 坐标/格子 → 下次用名字点
找字点击     先读 UIA 控件树，读不到再 OCR
屏幕点阵     64px 格子，坐标能说成 M8+12-6
文档转换     docx→pdf / xlsx→csv（走 LibreOffice / Excel COM）
截图取证     每次点击前后各截一张，写进日志
```

## 要求

```
Windows 10/11
Node 18+
.NET Framework 4（Windows 自带；只有重新编译 C# 工具时需要）
可选：LibreOffice（docx→pdf 才需要）、ffmpeg（截图放大/OCR 预处理）
```

## 安装

```bash
npm i -g dsh-desktop-toolkit
dsh-desktop list
```

或者直接用源码：

```bash
git clone <repo>
cd dsh-desktop-toolkit
npm run build        # 重新编译 bin/*.exe（已预编译也可以跳过）
node control.mjs list
```

## 常用命令

```bash
# 看窗口
node control.mjs window list
node control.mjs window lock 记事本      # 锁定；之后窗口外的坐标会被拒绝
node control.mjs window unlock

# 点坐标 / 拖拽
node control.mjs click 800 480
node control.mjs drag 400 300 900 600

# 点文字（先 UIA，后 OCR）
node control.mjs clickname "刷新"
node control.mjs clickname "确定" --index 2      # 同名多个时选第 2 个

# 看某个窗口的控件树
node control.mjs uia list <pid>
node control.mjs uia find <pid> "文件"

# 点阵
node control.mjs grid info
node control.mjs grid cell M8              # → (800,480)
node control.mjs grid cell "M8+12-6"       # 亚格偏移 → (812,474)

# OCR（常驻服务，空闲 10 分钟自己退）
node control.mjs ocr <图片> [要找的字]

# 别的一次性工具
node control.mjs docx2pdf a.docx a.pdf
node control.mjs xlsx2csv a.xlsx a.csv
node control.mjs frame video.mp4 12 out.jpg
```

## 设计原则

1. **先选通道，再执行，最后验证**：UIA 能读就用 UIA，读不到退 OCR，都没中就报错而不是瞎点
2. **动作要有证据**：每次操作写 `数据/操控日志.jsonl`，点击前后各一张截图
3. **窗口外不点**：`window lock` 之后坐标守卫会拦下越界操作
4. **如实报限制**：UWP 应用（如 Win11 新版记事本）读不到控件树，这是已知限制，见 KNOWN-LIMITS.md

## 已知限制

见 [KNOWN-LIMITS.md](KNOWN-LIMITS.md)。简单说：

```
UWP 应用 UIA 可能读不到 → 退回 OCR
OCR 会把形近字认错（本/不、末/未）→ 搜索词要短
最小化窗口的矩形是 -32000 假坐标 → lock 会拒绝
```

## 目录

```
control.mjs      主 CLI
ocr-svc.mjs      OCR 常驻服务客户端
bin/*.exe        编译好的 C# 小工具（mouse/grid/imgdiff/window/uia）
src/*.cs         对应源码，build.ps1 可重新编译
ps/*.ps1         截图、剪贴板输入、Office 转换等 PowerShell 脚本
```

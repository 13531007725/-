# Cell_ppt 运行原理、PowerPoint 自动化与自有 API 复刻

## 1. Cell_ppt 是什么

Cell_ppt 不是一个常驻运行的 PowerPoint 插件，而是一个 Codex Skill。它由 `SKILL.md`、参考说明和可执行脚本组成：

- `SKILL.md`：定义何时触发、必须遵循的工作流和质量要求。
- `scripts/`：负责图片处理、SVG 解析、路径缓存、API 通信和 PowerPoint 自动化。
- `references/`：定义后端、平台兼容性和图形转换约束。

当用户明确调用 Cell_ppt，或请求“把图片重建为 PowerPoint 原生可编辑对象”时，Codex 会读取技能说明并按其中的固定流程执行脚本。

## 2. 完整运行流程

```text
用户上传原图
   ↓
Codex 记录所有文字的位置、字号、颜色和绘制顺序
   ↓
Image 2 仅删除原图文字，保留箭头、线条和科学图形
   ↓
路径返回 API：无文字位图 → SVG 贝塞尔路径
   ↓
把记录的文字重新合并为 SVG <text>
   ↓
Python 将 SVG 解析为 geometry-cache.json
   ↓
删除完全重复的绘图路径，保留其他路径及原始绘制顺序
   ↓
Windows PowerShell 通过 COM 连接 PowerPoint
   ↓
BuildFreeform 创建原生自由曲线，AddTextbox 创建原生文本框
   ↓
保存并重新验证可编辑 PPTX
```

其核心调度逻辑可以概括为：

```powershell
& $vectorizer @vectorizerArgs
& $pythonExe $textMerger ...
& $pythonExe $cacheBuilder ...
& $pythonExe $visibilityCuller ...
& $pptRuntime @runtimeArgs
```

API 只负责生成 SVG 路径；真正操作 PowerPoint 的是本地脚本。

## 3. 为什么能够调用 PowerPoint

Windows 桌面版 PowerPoint 提供 COM Automation 对象模型，并注册了 `PowerPoint.Application` ProgID。PowerShell 可以直接创建或取得该 COM 对象：

```powershell
$application = New-Object -ComObject "PowerPoint.Application"
$application.Visible = -1

$presentation = $application.ActivePresentation
$slide = $application.ActiveWindow.View.Slide
```

创建原生文字框：

```powershell
$shape = $slide.Shapes.AddTextbox(
    1,
    [single]$left,
    [single]$top,
    [single]$width,
    [single]$height
)

$shape.TextFrame2.TextRange.Text = $text
$shape.TextFrame2.TextRange.Font.Name = $fontName
$shape.TextFrame2.TextRange.Font.Size = $fontSize
```

创建原生贝塞尔自由曲线：

```powershell
$builder = $slide.Shapes.BuildFreeform(
    1,
    [single]$startX,
    [single]$startY
)

$builder.AddNodes(
    1,
    1,
    [single]$control1X,
    [single]$control1Y,
    [single]$control2X,
    [single]$control2Y,
    [single]$endX,
    [single]$endY
)

$shape = $builder.ConvertToShape()
```

最后保存为标准 PPTX：

```powershell
$presentation.SaveAs($outputPath, 24)
```

因此输出不是一张不可编辑的整图，而是由文本框、自由曲线、填充和描边组成的 PowerPoint 原生对象。

## 4. 是否需要 pywin32

当前 Windows 版本的 Cell_ppt 不依赖 `pywin32`。它采用：

- PowerShell COM：控制 PowerPoint。
- Python：解析 SVG、计算贝塞尔节点、建立缓存和路径去重。
- HTTP 客户端：上传图片并下载 SVG。

如果希望全部使用 Python 重写 PowerPoint 控制层，可以选择 `pywin32`：

```python
import win32com.client

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True

presentation = ppt.ActivePresentation
slide = ppt.ActiveWindow.View.Slide

builder = slide.Shapes.BuildFreeform(1, 100, 100)
builder.AddNodes(0, 1, 200, 100)
builder.AddNodes(1, 1, 220, 100, 280, 180, 300, 200)

shape = builder.ConvertToShape()
shape.Line.Weight = 1.5
presentation.Save()
```

PowerShell COM 与 `pywin32` 调用的是同一个 PowerPoint 对象模型，只是编程语言不同。

## 5. 第三方 API 是否影响 PPT 作图质量

会。第三方路径 API 是决定几何精度的主要环节之一。

| 环节 | 主要影响 |
|---|---|
| 文字识别 | 字体、字号、上下标、颜色和文字位置 |
| 去文字处理 | 是否误删线条，是否改变背景和图形 |
| 路径 API | 曲线准确度、细节、颜色、节点数量和图层顺序 |
| SVG 解析 | 变换矩阵、贝塞尔控制点和复合路径是否正确 |
| PowerPoint COM | 是否忠实地把缓存转换为原生 Shape |

API 生成的 SVG 如果存在问题，PowerPoint 只会忠实保留这些问题，例如：

- 圆形变成不规则曲线；
- 细线、虚线或箭头丢失；
- 化学键连接不正确；
- 节点数量过多，导致 PPT 很卡；
- 填色、描边或图层顺序错误；
- 原本相同的颜色被近似成多个颜色。

PowerPoint 转换阶段不会自动修复语义错误。最终结果的几何质量上限基本由输入清理质量和 SVG 路径质量决定。

为提高可复现性，建议：

1. 固定 API 或模型版本，不使用随时变化的未版本化端点。
2. 保存 API 返回的原始 SVG，出现问题时可直接复现。
3. 建立一组标准测试图片，比较路径数量、颜色、viewBox 和最终 PPT 对象数量。
4. 文字始终单独记录并恢复，不让图像矢量化模型生成文字。
5. 对 API 输出执行 SVG 验证，拒绝栅格包装、外部资源和不兼容结构。

## 6. 替换成自己的 API

PowerPoint 绘图层和图片矢量化 API 是解耦的。只要自己的系统能提供合格的 SVG，就可以保留后续的文字合并、路径缓存和 PowerPoint 绘图代码。

推荐的 SVG 输出要求：

- 存在合法 `viewBox`；
- 只包含真实矢量路径和文字；
- 不包含 `<image>` 栅格节点；
- 使用直线或三次贝塞尔曲线；
- 使用纯色填充和纯色描边；
- 保留 SVG 从底层到顶层的字面绘制顺序；
- 不使用无法转换的 mask、clipPath、pattern 和外部链接；
- 元素 ID 唯一；
- 输出编码为 UTF-8。

### 方案 A：直接调用自己的已有 API

如果已有接口能同步返回 SVG，只需编写一个适配器：

```python
from pathlib import Path
import requests


def vectorize(input_image: str, output_svg: str, token: str) -> None:
    with open(input_image, "rb") as image:
        response = requests.post(
            "https://your-api.example.com/vectorize",
            headers={"Authorization": f"Bearer {token}"},
            files={"image": image},
            timeout=300,
        )

    response.raise_for_status()
    payload = response.json()
    Path(output_svg).write_text(payload["svg"], encoding="utf-8")
```

然后让主流程调用这个适配器，替换原来的路径 API 客户端。

### 方案 B：复刻异步 REST 协议

可以提供三个端点：

```http
POST /api/images
Content-Type: multipart/form-data

返回：
{"image_id": "abc123"}
```

```http
GET /api/images/abc123

处理中：
{"status": "processing"}

完成：
{"status": "completed"}
```

```http
GET /api/images/abc123/file
Content-Type: image/svg+xml

<svg ...>...</svg>
```

一个简化的 FastAPI 外壳如下：

```python
from fastapi import BackgroundTasks, FastAPI, HTTPException, UploadFile
from fastapi.responses import Response
from uuid import uuid4

app = FastAPI()
jobs = {}


def run_vectorizer(job_id: str, image_bytes: bytes) -> None:
    try:
        # 将这里替换为自己的模型或矢量化算法。
        svg_bytes = your_vectorizer(image_bytes)
        jobs[job_id] = {"status": "completed", "svg": svg_bytes}
    except Exception as exc:
        jobs[job_id] = {"status": "failed", "error": str(exc)}


@app.post("/api/images")
async def upload(image: UploadFile, tasks: BackgroundTasks):
    job_id = uuid4().hex
    image_bytes = await image.read()
    jobs[job_id] = {"status": "processing"}
    tasks.add_task(run_vectorizer, job_id, image_bytes)
    return {"image_id": job_id}


@app.get("/api/images/{job_id}")
def status(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404)
    return {"status": job["status"]}


@app.get("/api/images/{job_id}/file")
def download(job_id: str):
    job = jobs.get(job_id)
    if not job:
        raise HTTPException(404)
    if job["status"] != "completed":
        raise HTTPException(409)
    return Response(job["svg"], media_type="image/svg+xml")
```

这只是协议外壳。生产环境还需要：

- Bearer Token 或其他认证；
- HTTPS；
- 数据库或对象存储；
- Redis/Celery 等任务队列；
- 文件大小和类型限制；
- 超时、重试和并发控制；
- SVG 安全校验；
- 日志脱敏和图片自动删除策略。

## 7. 是否必须搭服务器

不一定。

### 个人单机使用

推荐本地适配器：

```text
PNG/JPEG → 本地 Python/模型 → SVG → Cell_ppt 后处理 → PowerPoint
```

优点是隐私好、部署简单、不需要网络服务。只要本地算法能输出质量合格的 SVG，就不需要服务器。

### 直接调用模型厂商

客户端可以直接调用厂商 API，也不需要自己的服务器。但要安全保存 API Key，并注意图片会离开本机。

### 多人或商业化使用

建议搭建自己的服务器，以便：

- 统一管理上游 API Key；
- 管理用户、额度和计费；
- 运行 GPU 矢量化模型；
- 执行队列、限流和重试；
- 统一升级模型版本；
- 记录质量指标和失败原因。

## 8. 推荐复刻路线

如果目的是个人研究或内部使用，最稳妥的路线是：

1. 保留现有 SVG 解析、几何缓存和 PowerPoint COM 绘图模块。
2. 新建自己的 `vectorize` 适配器，只替换“位图转 SVG”部分。
3. 先使用本地同步调用，不搭服务器。
4. 用 10–20 张科学示意图建立质量测试集。
5. 当需要多人使用、远程 GPU 或统一认证时，再把适配器封装成 FastAPI 服务。

这样无需重写整个 PowerPoint 系统，也能独立控制矢量化模型和数据隐私。


# 发票整理台

本地运行的 PDF 发票与收据整理工具。上传文件夹、多个 PDF 或 ZIP，自动识别国内发票与海外收据，检查购买方信息和重复号码，并下载分类 PDF、Excel 汇总与错误报告。

## 功能

- 国内发票按机票、高铁票、餐饮、打车、住宿五类归档；检查购买方名称、统一社会信用代码/纳税人识别号、发票号码及价税合计。
- 海外收据提取总金额和币种，按币种分别汇总；不同货币不会直接相加。
- 同号发票及内容完全相同的文件进入错误区，不计入有效金额。
- 优先读取 PDF 文字层；扫描件使用 OCR。分类规则无法判断时，可选择调用自己配置的大模型服务。
- 输出 ZIP 含国内总览及五类明细工作表、各类别的独立 Excel、海外汇总和错误报告。

## 环境要求

- Python 3.12
- [uv](https://docs.astral.sh/uv/getting-started/installation/)（推荐用于创建虚拟环境和安装依赖）
- macOS、Linux 或 Windows；有现代浏览器即可使用页面
- 首次安装 Python 依赖需要联网；日常识别不需要网络，除非主动启用大模型辅助分类

OCR 模型和 Python 依赖由 `requirements.txt` 安装。项目不要求单独安装 Tesseract。

## 安装与启动

克隆仓库后进入项目目录：

```bash
git clone https://github.com/ZhenfengSu/invoice-organizer.git
cd invoice-organizer
```

macOS 可直接运行：

```bash
./start.command
```

首次启动会创建 `.venv` 并安装依赖。默认只监听本机 `127.0.0.1:8765`，浏览器打开 [http://127.0.0.1:8765](http://127.0.0.1:8765)。需要换端口时：

```bash
PORT=9000 ./start.command
```

Linux 或希望手动启动时：

```bash
uv venv --python 3.12 .venv
uv pip install --python .venv/bin/python -r requirements.txt
.venv/bin/python -m uvicorn app:app --host 127.0.0.1 --port 8765
```

Windows PowerShell：

```powershell
uv venv --python 3.12 .venv
uv pip install --python .venv/Scripts/python.exe -r requirements.txt
.\.venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8765
```

更新代码后，重新运行安装依赖的命令再启动服务。停止服务按 `Ctrl+C`。

## 可选：大模型辅助分类

默认不需要 API。只有勾选页面中的“大模型辅助分类”，且普通规则无法确定类别时，程序才会调用所配置的服务。金额、税号和发票号码始终由本地代码提取和检查。

如需使用兼容 OpenAI Responses API 的服务，复制示例配置并在本机填写：

```bash
cp .env.example .env
```

`.env` 中设置 `LLM_BASE_URL`、`LLM_API_KEY` 和 `LLM_MODEL`。模型默认值为 `gpt-5.6-luna`；接口地址和密钥没有写入仓库。`.env` 已被 Git 忽略，不要把真实密钥、票据文本或生成结果提交到公开仓库。启用此功能会把待分类票据的文字发送给你配置的服务。

Windows 可以手动复制 `.env.example` 为 `.env`。修改 `.env` 后重启服务。

## 使用方法

1. 选择文件夹、多份 PDF，或包含 PDF 的 ZIP。文件夹和 ZIP 内的子目录会递归处理。
2. 输入允许的购买方名称及对应税号。多个主体按相同行号填写；留空则只检查字段能否识别及 18 位统一社会信用代码校验位。
3. 点击“开始整理”，检查页面摘要并下载结果 ZIP。

ZIP 内容：

```text
国内发票/
  国内发票汇总.xlsx       # 总览 + 五个类别工作表
  机票/ 高铁票/ 餐饮/ 打车/ 住宿/
    原始 PDF 与类别统计.xlsx
海外收据/
  海外收据汇总.xlsx       # 文件名、金额、币种及各币种小计
  原始 PDF
错误发票/
  错误报告.xlsx
  错误报告.txt
  待人工检查的 PDF
```

输出 PDF 保留原文件名；重名时追加序号。错误票据不会计入有效金额。遇到只写 `$` 而没有明确币种的海外收据，会进入错误区，避免错误合并货币。

## 校验与限制

- 18 位统一社会信用代码按 GB 32100 校验位算法检查；旧式 15 位税号只做字段和配置值匹配。
- 购买方“正确”指与页面填写的允许值一致，不代表税务平台验真，也不能判断票据是否真实开具或已报销。
- 仅写有“旅行社 + 退票费”且没有铁路信息的发票暂归为机票；这是基于票面文本的分类推断。
- 当前支持每份 1–10 页、每份最多 30 MB、单次最多 1000 份 PDF；上传或解压后的 PDF 总量最多 200 MB。
- 处理结果保存在本机临时目录，24 小时后在下一次请求时清理；服务重启后旧下载链接失效。

## 测试

安装依赖后运行：

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Windows 将 Python 路径替换为 `.venv\Scripts\python.exe`。测试只用合成数据，不包含真实发票或 API 凭据。

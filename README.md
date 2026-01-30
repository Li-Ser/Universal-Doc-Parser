# 🌌 Universal Doc Parser (通用智能文档解析器)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python](https://img.shields.io/badge/Python-3.8%2B-blue)](https://www.python.org/)

一个专为 **RAG (检索增强生成)** 知识库构建设计的智能文档分块与清洗工具。

它能自动识别文档结构（目录树、正则编号、视觉字号、Word样式），精准提取正文、表格和层级关系，并支持多格式导出。

## ✨ 核心特性

- **🧠 智能侦察路由**：
  - 自动识别 PDF 是否包含完备书签（Outlines）。
  - 自动检测正文是否为序列化编号（1.1, 1.2...）。
  - 针对无结构文档，自动启用视觉字号分析算法。
  - 原生支持 `.docx` Word 样式解析。

- **🔧 深度清洗与修复**：
  - **表格重构**：将 PDF 中的表格自动转换为 Markdown 格式嵌入正文，保留参数对应关系。
  - **语义修复**：自动合并被排版截断的单词（Hyphenation）和断句。
  - **去噪**：内置通用页眉、页脚、版权信息过滤库。

- **📝 上下文注入**：
  - 在分块内容中自动注入父级标题路径（如 `配置 > 接口 > 以太网接口`），提升向量检索准确度。

- **💾 全能导出**：
  - 支持 `Excel`, `JSON` (机器友好), `Word`, `PDF` 四种格式输出。

## 🚀 快速开始

### 1. 安装依赖
```bash
git clone [https://github.com/你的用户名/Universal-Doc-Parser.git](https://github.com/你的用户名/Universal-Doc-Parser.git)
cd Universal-Doc-Parser
pip install -r requirements.txt
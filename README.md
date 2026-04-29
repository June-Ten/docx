# DOCX 划词批注

前端：`frontend`，Vue 3 + Vite + mammoth.js + web-highlighter。

后端：`backend-python`，Flask + python-docx。

## 运行

```bash
cd backend-python
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

```bash
cd frontend
npm install
npm run dev
```

默认地址：

- 前端：http://localhost:5173
- 后端：http://localhost:5000

## 功能

- 前端上传 DOCX 后用 mammoth.js 转成 HTML 预览。
- 在预览区域划词后，web-highlighter 自动高亮选区。
- 保存批注后，点击“生成批注文件”会把原始 DOCX 和批注列表发送到 Python 后端。
- 后端用 python-docx 的 `Document.add_comment()` 在匹配文本上生成 Word 批注。

## 说明

当前版本按 `选中文本 + 第几次出现` 定位原文。建议批注单段落内的文本；跨段落选择、复杂表格/页眉页脚内容可能需要进一步增强定位逻辑。

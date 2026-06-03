# 考试批改助手

基于 Streamlit 的试卷手动批改工具。

## 功能

- 上传参考答案 .docx 和学生答卷 .docx（批量）
- 逐人逐题对照批改，每题含"程序"(3分) + "结果"(2分)
- 分数自动保存，关闭浏览器不丢失
- 导出 .xlsx 成绩表，含每题明细分和总分
- 导出前自动检查未批改题目并提醒

## 安装

```bash
pip install -r requirements.txt
```

## 使用

```bash
streamlit run grading_app.py
```

## 文件结构

```
grading-assistant/
├── grading_app.py     # Streamlit 主程序
├── docx_parser.py      # .docx 解析（封面信息 + 题目结构 + 图片提取）
├── scoring.py          # 计分 + .xlsx 导出
├── requirements.txt    # Python 依赖
├── output/             # 自动保存 & 导出
└── README.md
```

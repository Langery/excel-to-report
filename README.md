# excel-to-report（表格成文助手）

把一张 Excel/CSV 表格变成**明天早上能直接交出去的文字材料**。不做泛泛的"数据分析"，只做职场人真正会交付的四类东西。

| 交付物 | 场景 |
|---|---|
| 周报初稿 | 本周工作量化呈现（做了多少、进展到哪、下周计划） |
| 工作汇报段落 | 给领导的结论先行式短文，可直接粘进邮件 / IM |
| 分类叙述 | 按部门 / 月份 / 状态等维度划分的成段文字 |
| 异常项清单 | 表格里"不对劲"的数据点名（超期、空值、离群值） |

## 核心纪律

- **绝不编造数字**。报告中每个数字都必须来自表格本身；表里没有的信息写"表中未提供"。
- **结论先行**。第一句就是最重要的结论，后面才是数据支撑。
- **口径先确认**。动笔前问清时间范围、汇报对象、划分维度。口径不同，同一张表的写法完全不同。

## 安装

### Claude Code

```bash
git clone https://github.com/Langery/excel-to-report.git ~/.claude/skills/excel-to-report
```

### WorkBuddy / OpenClaw 系变体

```bash
git clone https://github.com/Langery/excel-to-report.git ~/.workbuddy/skills/excel-to-report
```

### SkillHub 一键安装

```bash
skillhub install excel-to-report --dir~/.workbuddy/skills/
```

## 使用

安装后在支持 Skill 的 Agent 里说人话即可触发，例如：

- "帮我把这个表汇总成周报"
- "按部门分类写成汇报段落"
- "这个表里哪些数据不对劲？"

也可以直接命令行跑：

```bash
# 读表 + 类型推断
python scripts/read_table.py 表格.xlsx

# 按维度分组统计（日期列自动按月聚合）
python scripts/read_table.py 表格.xlsx --group 创建日期
```

## 目录结构

```
excel-to-report/
├── SKILL.md                        # 入口，六步工作流
├── scripts/
│   └── read_table.py               # 读表 + 类型推断 + 分组统计，输出 JSON
├── references/
│   ├── output-contract.md          # 读表 JSON 字段契约（写数字前必读）
│   ├── report-types.md             # 四类材料的适用场景与结构
│   └── writing-style.md            # 写作风格规范
├── assets/templates/
│   ├── weekly-report.md
│   ├── briefing.md
│   ├── category-narrative.md
│   └── anomaly-checklist.md
└── requirements.txt
```

`read_table.py` 的 JSON 输出中，部分数值是脚本已算好的汇总值，部分需要自己从行数据推算——**动笔写数字前先读 `references/output-contract.md`**，这是本 Skill 防止编造数字的关键防线。

## 依赖

```bash
pip install -r requirements.txt   # xlsx 读取需要 openpyxl；CSV 走标准库无依赖
```

xlsx 环境缺失时脚本会返回明确提示，此时可改用 Agent 内置表格读取能力，或请用户提供 CSV 版本。

## 不在范围

不做图表、不修改原始表格（只读）、不做数据清洗、不预测表格里没有的结论。

## License

MIT
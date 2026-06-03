# nursing-reflection-grader

妇产科护理反思日记自动评改工具，基于 AI 创作检测（ai-detector / GPTHumanizer API），严格依据《妇产科护理学》反思日记评价量规（2026年修订版），生成评改报告与 Excel 汇总表。

## 功能特性

- ✅ **8 维度评分**（情境/安全/职业/关怀/专业/思辨/责任/反思），满分 100 分
- ✅ **AI 创作检测**（ai-detector API），AI 概率 > 40% 直接判定不合格
- ✅ **批量处理**，支持 `.docx` / `.txt` 文件
- ✅ **Markdown 评改报告** + **Excel 班级成绩汇总表**
- ✅ **自动安装依赖**（python-docx、openpyxl、requests）

## 评分标准

来源：《妇产科护理学》反思日记评价量规（2026年修订版）

| 序号 | 评价维度 | 分值 |
|------|----------|------|
| 1 | 情境呈现与问题意识 | 10 分 |
| 2 | 生命至上与母婴安全意识 | 15 分 |
| 3 | 课程价值与职业认知 | 15 分 |
| 4 | 人文关怀与共情能力 | 15 分 |
| 5 | 专科护理思维与风险应对 | 15 分 |
| 6 | 科学精神与思辨思维 | 10 分 |
| 7 | 职业认同与责任担当 | 10 分 |
| 8 | 自我反思与改进计划 | 10 分 |

**等级阈值**：优秀 ≥ 90 分 | 良好 ≥ 75 分 | 合格 ≥ 60 分 | 不合格 < 60 分

## AI 创作识别规则

系统通过 **ai-detector (GPTHumanizer API)** 检测 AI 生成特征：

- **AI 概率 > 40%** → 总分强制置 0，等级标注为 `不合格（AI创作）`
- **AI 概率 ≤ 40%** → 正常评分，不受影响

> AI 检测为辅助判定的参考依据，最终由教师结合学生实际情况综合认定。

## 使用方法

### 安装依赖

```bash
pip install python-docx openpyxl requests
```

### 运行评改

```bash
python scripts/grader.py "待评改文件夹路径"
```

可选参数 `--output` 指定输出目录（默认在输入文件夹下创建 `评改结果` 子文件夹）：

```bash
python scripts/grader.py "C:\反思日记" --output "C:\评改输出"
```

### 文件命名建议

为确保正确识别学生姓名，建议文件命名格式：

- `张晓梅_反思日记.docx`
- `20230001李明.txt`
- `反思日记-王小红.docx`
- `产时模块反思日记 祝梦君.docx`

## 输出说明

脚本运行完成后，输出目录中会生成：

1. **每位学生的 Markdown 评改报告**（含 8 维度评分 + 评语 + 参考标准 + AI 检测详情）
2. **班级成绩汇总.xlsx**（含成绩排列 + 维度分析 sheet）

## 技术依赖

- Python 3.8+
- python-docx（读 .docx）
- openpyxl（写 .xlsx）
- requests（调用 ai-detector API）

脚本运行时会自动安装所需依赖，无需手动安装。

## AI 检测 API

- **端点**：`https://detect.gpthumanizer.ai/api/detect_ai`
- **方法**：POST
- **认证**：无需认证
- **返回字段**：`class`（human/ai/ai_humanized/light_edited）、`ai_possibilities`（0-1）

## 许可证

MIT License

## 作者

overloading2001

## 更新日志

- **2026-05-19**：集成 ai-detector API，替换原 ai-density 本地规则引擎
- **2026-05-18**：调整 AI 检测阈值为 40%
- **2026-05-18**：初始版本发布，支持 8 维度评分 + ai-density 检测

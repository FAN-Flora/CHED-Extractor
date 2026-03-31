import os
import re
import json
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openai import OpenAI
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

# =================================================================
# 1. 配置区
# =================================================================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "Replace this with your API-KEY")

INPUT_DOCX = r"\CHED\中国三千年疫灾史料汇编.docx" # Replace this with your historical material text path
OUTPUT_EXCEL = r"\CHED\Epidemic_record.xlsx"  # Replace this with your result output path

# 系统 Prompt（在原有完整内容基础上，添加了禁止翻译的约束）
SYSTEM_PROMPT = """
# 角色：
中国古代疫灾分析智能体

你是一个专门处理中国古代疫灾史料的智能分析系统，负责从原始文献中精准提取并结构化疫灾事件的六大核心要素：疫时、疫域、疫因、疫果、疫种、疫策。你的输出将用于构建可共享、可验证、可复现的学术数据集，因此必须严格遵循格式规范与语义准确性。

⚠️ **输出格式重要提示**：请始终以 JSON 格式返回结果，JSON 对象必须包含以下字段，不要添加任何额外文字。

## 目标：
对输入的古代疫灾史料文本进行逐段解析，识别以下信息，并生成符合指定格式的 Excel 表格：
- 疫时（公历年份、季节、农历月份）
- 疫域（历史与现代行政区划）
- 疫因（导致疫灾的原因）
- 疫果（疫灾造成的后果）
- 疫种（疫病种类）
- 疫策（应对措施）

⚠️ 特别注意：**“饥荒”“人相食”等表述可能为疫因或疫果，必须依据上下文语义判断**。例如：
> 示例1：  
> 原文：“饥，人相食，饿莩盈野，瘟疫大作。”  
> 分析：此处“人相食”发生在“瘟疫大作”之前，是社会崩溃的表现，**作为疫灾的诱因**。  
> → 疫因 = “饥、人相食”（关键词）  
> → 疫果 = （留空）  

> 示例2：  
> 原文：“大疫，死者过半，民无所食，人相食。”  
> 分析：此处“人相食”是疫灾导致的社会后果。  
> → 疫果 = “大疫，死者过半，民无所食，人相食。”（完整句）  
> → 疫因 = （留空或填写其他原因）

## 输出格式（必须严格遵守）：
生成一个名为 `中国古代疫灾史料解析结果.xlsx` 的 Excel 文件，包含以下 **15 列**（顺序不可更改）：

| 列名 | 说明 |
|------|------|
| 代码 | 自增序号，从1开始 |
| 年份 | 公历年份（公元前用负数，如 -771），无则留空 |
| 季节 | 春、夏、秋、冬等，按原文填写 |
| 月份（农历） | 如“五月”“十月”，按原文填写 |
| 原文行政区划记录 | 史料中的原始地名（如“镐”“汴梁”） |
| 一级区划 | 历史高层政区（如州、道、路） |
| 二级区划 | 历史中层政区（如府、郡） |
| 县 | 历史县级单位 |
| 省份 | 对应现代省级行政区 |
| 县名 | 对应现代县级行政区 |
| 疫种 | 疫病种类关键词，多个用顿号分隔，无则留空 |
| 疫因 | 导致疫灾的原因关键词，多个用顿号分隔，无则留空 |
| 疫果 | ✅ **包含疫灾后果的完整原始语句（保留上下文）**，无则留空 |
| 疫策 | ✅ **包含应对措施的完整原始语句（保留上下文）**，无则留空 |
| 备注 | 可填写推断依据、歧义说明、史料出处等 |

### 字段填写规则：
- **“疫果”和“疫策”必须填写完整句子**，不得仅写关键词；  
- **“疫因”和“疫种”以关键词为主**，多个用顿号分隔；  
- 所有文本使用简体中文，忠实于原文；  
- 无信息字段**留空**，不填“无”“—”或“……”。

## 重要约束（必须遵守）：
- 所有提取结果必须严格采用史料原文中的原始表述，**不得翻译成现代汉语**，不得进行改写或解释。
- **疫因**：直接摘录原文中导致疫灾的原因的关键词（如“旱”“饥”“人相食”），多个用顿号分隔，保留原文用词。
- **疫果**：直接复制原文中包含疫灾后果的完整句子（如“大疫，死者相枕”），不得改写。
- **疫种**：直接摘录原文中疫病种类的关键词（如“瘟疫”“痢疾”），保留原文用词。
- **疫策**：直接复制原文中包含应对措施的完整句子（如“官设粥厂，瘗尸掩骼”），不得改写。
- **备注**：如需说明推断依据，可使用现代汉语，但疫因/疫果/疫种/疫策仍须保持原文。

## 工作流：
1. 按段落拆分输入文本，每段通常对应一个县域的疫灾记录。
2. 若段首出现省/府名（如“河南”“关东”），则其后所有记录归属该政区，直至下一个省级单位出现。
3. 对每段执行：
   - 提取时间信息，转换为公历年份；
   - 识别疫域，结合《中国历史地图集》《中国行政区划通史》补全历史与现代区划；
   - 分析语句语义，判断模糊项（如“人相食”）属于疫因还是疫果；
   - 提取“疫果”“疫策”时，直接复制包含相关语义的**完整句子**；
   - 提取“疫因”“疫种”时，从下述关键词库中匹配并提取。

## 术语分类参考（用于关键词提取）

### 1. 疫时（Time）
- 包括：年份（需转公历）、年号、季节（春、夏、秋、冬）、农历月份（正月、二月…腊月）。
- 示例：“崇祯七年春三月” → 年份=1634，季节=春，月份（农历）=三月。

### 2. 疫域（Place）
- 原文地名需保留于“原文行政区划记录”；
- 历史区划（一级/二级/县）依据《中国历史地图集》补全；
- 现代区划（省份/县名）依据当前行政区划标准匹配。

### 3. 疫因（Cause of Epidemic）
以下词汇若在“瘟疫发生前”出现，视为疫因关键词：
旱、水、潮水、涝、雨、淫雨、黑雨、黑眚、白眚、饥、翻坛、火、饥荒、歉收、地震、震、蝗虫、螟虫、星异、鸟异、兵、乱、战争、工役、劳役、修筑、难民集聚、米运传播、传入、带来、人相食

> 注：“人相食”仅当出现在疫灾描述之前或作为背景条件时，才归为疫因。

### 4. 疫果（Effect of Epidemic）
以下词汇若在“瘟疫发生后”出现，其所在**完整句子**应填入“疫果”列：
疫死、死绝、死者、死亡、人多死、人民凋敝、死绝、流殍、病毙、流离、流亡、迁徙、瘟疫盛行、几停贸易、歉收、粮价上涨、粮价飞涨、米贵、米价腾贵、米价翔贵、米价骤昂、废耕、棺材售罄、人相食、店民惊骇、人民惊骇、撤兵、引军还、兵溃、军撤、班师、疫苗接种

> 注：“人相食”若作为疫灾后果出现（如“疫甚，人相食”），则整句填入“疫果”。

### 5. 疫种（Category of Epidemic）
以下为疫病种类关键词，可多选，用顿号分隔：
鼠疫、天花、霍乱、猩红热、红汗疹、疟疾、斑疹伤寒、流行性脑脊髓膜炎、流脑、白喉、伤寒、副伤寒、回归热、痢疾、赤痢、菌痢、流感、黑热病、血吸虫、麻疹、流行性腮腺炎、大头瘟、炭疽病、马蹄瘟、百日咳、肺结核、结核病、出血热、麻风病、丝虫病、登革热、肺炎、脑炎、乙脑、嗜睡性脑炎、钩体病、结膜炎、脊髓灰质炎、肝炎、黄疸病、梅毒、布氏杆菌病、疟痢、痘疹、疥疮、沙眼、钩虫病、雅司病、克山病、大骨节病、地甲病、瘴气、瘴疫、热病、暑湿病、痧疫、痧、水痘、康花、汗疹、出水病、散疸症

### 6. 疫策（Response of Epidemic）
以下词汇若表示应对行为，其所在**完整句子**应填入“疫策”列：
过年、免、诏、谕、施、赈、恤、粥、禁遏、医、药、傩、祈、醮、禳、迎神、赛会、埋葬、救济、消毒、防疫、告示、接种、注射、预防、种痘、报告、避、焚、万人坑、禁屠、放假、瘟神、医院、医疗队

> 注：“迎...神”包括“迎城隍”“迎瘟神”等；“施”“赈”常与“粥”“药”“钱”搭配，需保留完整语境。

## 限制：
- 不得新增或删除列；
- 不得合并单元格或使用换行符；
- 必须确保输出表格可被 Excel 正常打开；
- 所有判断必须基于原文语义，不得主观臆断。

## 示例输入与输出：

🔹 输入段落：
“崇祯七年，夏，五月，陕西西安府大旱，饥，人相食，饿莩盈野，瘟疫大作。”

🔹 对应输出行：
| 代码 | 年份 | 季节 | 月份（农历） | 原文行政区划记录 | 一级区划 | 二级区划 | 县 | 省份 | 县名 | 疫种 | 疫因 | 疫果 | 疫策 | 备注 |
|------|------|------|--------------|------------------|----------|----------|----|------|------|------|------------------|------|------|----------------------------------|
| 1    | 1634 | 夏   | 五月         | 陕西西安府       | 陕西     | 西安府   |    | 陕西 | 西安市 |        | 旱、饥、人相食   |      |      | “人相食”在“瘟疫大作”前，判为疫因 |

🔹 输入段落：
“光绪二十六年，秋，直隶保定府大疫，死者相枕，官设粥厂，瘗尸掩骼。”

🔹 对应输出行：
| 代码 | 年份 | 季节 | 月份（农历） | 原文行政区划记录 | 一级区划 | 二级区划 | 县 | 省份 | 县名 | 疫种 | 疫因 | 疫果 | 疫策 | 备注 |
|------|------|------|--------------|------------------|----------|----------|----|------|------|------|------|------------------------------|------------------------------|------|
| 2    | 1900 | 秋   |              | 直隶保定府       | 直隶     | 保定府   |    | 河北 | 保定市 |        |      | “大疫，死者相枕”             | “官设粥厂，瘗尸掩骼”         |      |

---
请严格按照以上规则解析输入史料，并输出符合格式的 JSON 对象，包含以下字段：
{
  "年份": "",
  "季节": "",
  "月份（农历）": "",
  "原文行政区划记录": "",
  "一级区划": "",
  "二级区划": "",
  "县": "",
  "省份": "",
  "县名": "",
  "疫种": "",
  "疫因": "",
  "疫果": "",
  "疫策": "",
  "备注": ""
}
"""

# =================================================================
# 2. 逻辑解析区
# =================================================================

class AdvancedEpidemicAnalyzer:
    def __init__(self):
        if "替换" in DEEPSEEK_API_KEY or not DEEPSEEK_API_KEY:
            raise ValueError("错误：未检测到有效 API Key。")
        self.client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")

    def parse_document_with_years(self, path):
        """
        解析文档结构：识别年份标题（黑体四号居中），将后续正文段落归属到对应年份。
        返回列表，每个元素为 (year, paragraph_text)，其中 year 可能为 None（无年份归属）。
        """
        if not os.path.exists(path):
            raise FileNotFoundError(f"未找到文件: {path}")

        doc = Document(path)
        result = []
        current_year = None

        # 年份标题匹配正则：匹配 (AD1234) 或 (1234年) 等常见格式
        year_pattern = re.compile(r'[（(](?:AD)?(\d{4})(?:年)?[）)]')

        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            # 判断是否为年份标题（居中对齐且包含年份模式）
            is_title = False
            if p.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                match = year_pattern.search(text)
                if match:
                    year_num = int(match.group(1))
                    current_year = year_num
                    is_title = True

            if not is_title:
                if len(text) > 10 and not re.match(r'^\[\d+:\s', text):
                    result.append((current_year, text))

        return result

    def call_ai(self, year, text):
        """
        调用 AI，将年份信息作为上下文附加到用户消息中，并强调保留原文。
        """
        # 构建用户消息，强调不得翻译
        base_instruction = "请基于此解析以下段落，并以 JSON 格式返回结果。"
        no_translate = "注意：疫因、疫果、疫种、疫策必须直接使用原文中的词语和句子，不得翻译成现代汉语，不得改写。"
        if year is not None:
            user_content = f"该段史料所属的年份为 {year} 年（公历）。{base_instruction} {no_translate}\n\n段落内容：\n{text}"
        else:
            user_content = f"请解析以下疫灾史料段落（年份需从文本中推断）。{base_instruction} {no_translate}\n\n段落内容：\n{text}"

        try:
            response = self.client.chat.completions.create(
                model="deepseek-reasoner",
                messages=[
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": user_content}
                ],
                response_format={"type": "json_object"}
            )
            content = response.choices[0].message.content.strip()
            print(f"✅ AI 返回内容：{content}")  # 调试用
            return json.loads(content)
        except Exception as e:
            print(f"❌ AI 调用失败：{e}")
            return {"error": str(e), "raw_text": text, "year": year}

    def run(self):
        year_para_pairs = self.parse_document_with_years(INPUT_DOCX)
        print(f"📊 解析到 {len(year_para_pairs)} 个有效段落，开始调用 AI 处理...")

        final_records = []
        with ThreadPoolExecutor(max_workers=3) as executor:
            future_to_pair = {
                executor.submit(self.call_ai, year, text): (year, text)
                for year, text in year_para_pairs
            }

            for future in tqdm(as_completed(future_to_pair), total=len(year_para_pairs), desc="AI 解析进度"):
                res = future.result()
                if "error" in res:
                    continue
                if isinstance(res, dict):
                    # 如果 AI 返回的年份为空，但传入了年份，则用传入年份补充
                    if not res.get("年份") and res.get("year") is not None:
                        res["年份"] = res["year"]
                    record = [
                        len(final_records) + 1,
                        res.get("年份", ""),
                        res.get("季节", ""),
                        res.get("月份（农历）", ""),
                        res.get("原文行政区划记录", ""),
                        res.get("一级区划", ""),
                        res.get("二级区划", ""),
                        res.get("县", ""),
                        res.get("省份", ""),
                        res.get("县名", ""),
                        res.get("疫种", ""),
                        res.get("疫因", ""),
                        res.get("疫果", ""),
                        res.get("疫策", ""),
                        res.get("备注", "")
                    ]
                    final_records.append(record)

        cols = ["代码", "年份", "季节", "月份（农历）", "原文行政区划记录",
                "一级区划", "二级区划", "县", "省份", "县名",
                "疫种", "疫因", "疫果", "疫策", "备注"]
        df = pd.DataFrame(final_records, columns=cols)
        df.to_excel(OUTPUT_EXCEL, index=False)
        print(f"\n✅ 处理完成！结果已保存至: {OUTPUT_EXCEL}")


if __name__ == "__main__":
    analyzer = AdvancedEpidemicAnalyzer()
    analyzer.run()

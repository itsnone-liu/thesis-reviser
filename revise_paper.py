# -*- coding: utf-8 -*-
"""
论文修改程序 - 解析与诊断模块（revise_paper.py 前半部分）
"""

from docx import Document
import re
import json
import os
import time
import shutil
import uuid
import argparse
from web import (
    call_llm,
    build_chart_rule,
    MG_WORD_LIMITS,
    MG_SOFT_MAX,
    MG_CHAPTERS,
    txt_to_docx_safe,
    tasks_db,
    is_system_busy,
    system_lock,
)


# ================== [ 文档解析 ] ==================
def extract_docx_text(path: str) -> str:
    """使用 python-docx 读取 .docx 文件的所有段落文本，保留换行，过滤空段落"""
    try:
        doc = Document(path)
        paragraphs = [p.text.strip() for p in doc.paragraphs if p.text and p.text.strip()]
        return "\n".join(paragraphs)
    except Exception as e:
        print(f"读取文档失败 {path}: {e}")
        return ""


# ================== [ 论文画像分析 ] ==================
def analyze_original_paper(text: str) -> dict:
    """调用 LLM 分析论文全文，提取研究画像"""
    prompt = f"""你是一名学术论文分析专家。请对以下论文全文进行深度分析，输出严格JSON格式（不要markdown代码块，不要任何其他文字）。

论文全文：
{text[:8000]}

请输出以下字段的JSON：
{{
  "title": "论文标题",
  "major": "专业",
  "company": "研究对象/公司",
  "industry": "所属行业",
  "core_problems": ["问题1", "问题2", "问题3"],
  "data_hints": ["数据线索1", "线索2", "线索3"],
  "outline": {{
    "引言": "本章核心内容摘要（100字内）",
    "国内外研究现状": "...",
    "现状分析": "...",
    "问题与原因分析": "...",
    "解决方案": "...",
    "结论": "..."
  }},
  "theories": ["理论1", "理论2"],
  "problem_analysis_summary": "第三章核心逻辑",
  "solution_summary": "第五章方案概要",
  "charts_mentioned": ["图表1描述"],
  "data_quality_note": "数据质量评价"
}}

只输出JSON，严禁输出其他内容。"""

    for attempt in range(3):  # 首次 + 最多2次重试
        response = call_llm(prompt, max_tokens=3000)
        if not response:
            time.sleep(1)
            continue
        try:
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                # 补齐缺失字段
                if "title" not in result:
                    result["title"] = text[:50].replace('\n', ' ').strip() if text else "未识别标题"
                if "major" not in result:
                    result["major"] = "工商管理"
                if "company" not in result:
                    result["company"] = "某企业"
                if "industry" not in result:
                    result["industry"] = "制造业"
                if "core_problems" not in result:
                    result["core_problems"] = []
                if "data_hints" not in result:
                    result["data_hints"] = []
                if "outline" not in result:
                    result["outline"] = {name: "" for name, _ in MG_CHAPTERS}
                else:
                    for name, _ in MG_CHAPTERS:
                        if name not in result["outline"]:
                            result["outline"][name] = ""
                if "theories" not in result:
                    result["theories"] = []
                if "problem_analysis_summary" not in result:
                    result["problem_analysis_summary"] = ""
                if "solution_summary" not in result:
                    result["solution_summary"] = ""
                if "charts_mentioned" not in result:
                    result["charts_mentioned"] = []
                if "data_quality_note" not in result:
                    result["data_quality_note"] = ""
                return result
        except Exception as e:
            print(f"解析论文分析结果失败（尝试{attempt+1}/3）：{e}")
            time.sleep(1)
            continue

    # 兜底默认
    fallback_title = text[:100].replace('\n', ' ').strip() if text else "未识别标题"
    return {
        "title": fallback_title,
        "major": "工商管理",
        "company": "某企业",
        "industry": "制造业",
        "core_problems": ["问题待识别", "问题待识别", "问题待识别"],
        "data_hints": ["数据线索待识别", "数据线索待识别", "数据线索待识别"],
        "outline": {name: "" for name, _ in MG_CHAPTERS},
        "theories": ["理论待识别"],
        "problem_analysis_summary": "待分析",
        "solution_summary": "待分析",
        "charts_mentioned": [],
        "data_quality_note": "LLM解析失败，使用默认画像"
    }


# ================== [ 深度诊断与重构 ] ==================
def diagnose_and_reconstruct(original_profile: dict) -> dict:
    """调用 LLM 基于论文画像进行深度诊断，输出修订计划"""
    prompt = f"""你是一名资深学术论文评审与修改专家。请基于以下论文画像进行深度诊断，并输出修订方案。

论文画像：
{json.dumps(original_profile, ensure_ascii=False, indent=2)}

请输出严格JSON格式（不要markdown代码块，不要任何其他文字）：
{{
  "diagnosis_report": {{
    "overall_score": "1-10分",
    "overall_comment": "总体评价（200字）",
    "issues": [
      {{
        "chapter": "涉及章节",
        "severity": "高/中/低",
        "issue_type": "理论错误/数据矛盾/逻辑混乱/结构失衡/表述问题/其他",
        "description": "问题描述",
        "fix_direction": "修正方向"
      }}
    ]
  }},
  "revised_profile": {{
    "title": "优化后的标题",
    "company": "公司",
    "industry": "行业",
    "core_problems": ["修正后问题1", "问题2", "问题3"],
    "data_hints": ["数据方向1", "方向2", "方向3"]
  }},
  "chapter_revisions": {{
    "引言": {{"issues": [], "fix_requirements": [], "key_points": []}},
    "国内外研究现状": {{"issues": [], "fix_requirements": [], "key_points": []}},
    "现状分析": {{"issues": [], "fix_requirements": ["必须含至少2个<chart/>和1个<table/>"], "key_points": []}},
    "问题与原因分析": {{"issues": [], "fix_requirements": ["必须含至少1个<chart/>和1个<table/>"], "key_points": []}},
    "解决方案": {{"issues": [], "fix_requirements": ["必须含至少1个<chart/>"], "key_points": []}},
    "结论": {{"issues": [], "fix_requirements": [], "key_points": []}}
  }},
  "theory_adjustment": "理论框架调整说明",
  "data_strategy": "数据/图表策略",
  "special_notes": "生成新论文时的特殊注意事项"
}}

注意：chapter_revisions 中的每个章节都要有 issues、fix_requirements、key_points 三个字段。如果某章节无问题，请留空列表[]。
只输出JSON，严禁输出其他内容。"""

    for attempt in range(3):  # 首次 + 最多2次重试
        response = call_llm(prompt, max_tokens=4000)
        if not response:
            time.sleep(1)
            continue
        try:
            json_match = re.search(r'\{.*\}', response, re.DOTALL)
            if json_match:
                result = json.loads(json_match.group())
                # 补齐缺失字段
                if "diagnosis_report" not in result:
                    result["diagnosis_report"] = {"overall_score": "5", "overall_comment": "默认评价", "issues": []}
                if "revised_profile" not in result:
                    result["revised_profile"] = {
                        "title": original_profile.get("title", ""),
                        "company": original_profile.get("company", ""),
                        "industry": original_profile.get("industry", ""),
                        "core_problems": original_profile.get("core_problems", []),
                        "data_hints": original_profile.get("data_hints", [])
                    }
                if "chapter_revisions" not in result:
                    result["chapter_revisions"] = {}
                for name, _ in MG_CHAPTERS:
                    if name not in result["chapter_revisions"]:
                        result["chapter_revisions"][name] = {"issues": [], "fix_requirements": [], "key_points": []}
                    else:
                        for key in ["issues", "fix_requirements", "key_points"]:
                            if key not in result["chapter_revisions"][name]:
                                result["chapter_revisions"][name][key] = []
                if "theory_adjustment" not in result:
                    result["theory_adjustment"] = ""
                if "data_strategy" not in result:
                    result["data_strategy"] = ""
                if "special_notes" not in result:
                    result["special_notes"] = ""
                return result
        except Exception as e:
            print(f"解析诊断结果失败（尝试{attempt+1}/3）：{e}")
            time.sleep(1)
            continue

    # 保守兜底
    return {
        "diagnosis_report": {
            "overall_score": "5",
            "overall_comment": "由于LLM解析失败，采用保守默认诊断。建议人工复核论文内容。",
            "issues": []
        },
        "revised_profile": {
            "title": original_profile.get("title", ""),
            "company": original_profile.get("company", ""),
            "industry": original_profile.get("industry", ""),
            "core_problems": original_profile.get("core_problems", []),
            "data_hints": original_profile.get("data_hints", [])
        },
        "chapter_revisions": {
            name: {"issues": [], "fix_requirements": [], "key_points": []}
            for name, _ in MG_CHAPTERS
        },
        "theory_adjustment": "",
        "data_strategy": "",
        "special_notes": "LLM诊断解析失败，请人工检查原始论文并制定修改策略。"
    }


# ================== [ 诊断报告持久化 ] ==================
def save_diagnosis_report(revision_plan: dict, output_path: str):
    """将 revision_plan['diagnosis_report'] 格式化为易读文本报告并写入文件"""
    diag = revision_plan.get("diagnosis_report", {})
    lines = []
    lines.append("=" * 50)
    lines.append("【论文诊断报告】")
    lines.append("=" * 50)
    lines.append("")
    lines.append(f"总体评分：{diag.get('overall_score', 'N/A')}")
    lines.append("")
    lines.append("【总体评价】")
    lines.append(diag.get("overall_comment", "暂无"))
    lines.append("")
    issues = diag.get("issues", [])
    if issues:
        lines.append("【问题清单】")
        for i, issue in enumerate(issues, 1):
            lines.append(f"  问题{i}：")
            lines.append(f"    涉及章节：{issue.get('chapter', 'N/A')}")
            lines.append(f"    严重程度：{issue.get('severity', 'N/A')}")
            lines.append(f"    问题类型：{issue.get('issue_type', 'N/A')}")
            lines.append(f"    问题描述：{issue.get('description', 'N/A')}")
            lines.append(f"    修正方向：{issue.get('fix_direction', 'N/A')}")
            lines.append("")
    else:
        lines.append("【问题清单】")
        lines.append("  未发现明显问题（或LLM未返回问题列表）")
        lines.append("")
    lines.append("=" * 50)

    dir_name = os.path.dirname(output_path)
    if dir_name:
        os.makedirs(dir_name, exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))


# ================== [ 画像融合 ] ==================
def merge_into_standard_profile(revised_profile: dict, major: str) -> dict:
    """将修订画像转换为兼容原 web.py 的标准 profile 格式"""
    return {
        "major": major,
        "title": revised_profile.get("title", ""),
        "company": revised_profile.get("company", ""),
        "industry": revised_profile.get("industry", ""),
        "core_problems": revised_profile.get("core_problems", []),
        "data_hints": revised_profile.get("data_hints", []),
    }


# ================== [ 重写章节生成 ] ==================
def revise_gen_chapter(name, num, profile, chapter_rev, update):
    """修改版章节生成：完全复用原 gen_chapter 逻辑，追加重写要求"""
    update(f"正在生成第{num}章 {name}...", int((num - 1) / 6 * 80))
    if name == "引言":
        sp = "\n- 禁止使用任何标题、编号、markdown。用2-3段写背景和意义。"
    elif name == "国内外研究现状":
        sp = f"\n- 禁止使用任何标题、编号、markdown。纯段落叙述，不谈{profile['company']}。"
    elif name == "现状分析":
        sp = "\n- 小标题：3.1、3.2、3.3。必须含至少2个<chart/>和1个<table/>。"
    elif name == "问题与原因分析":
        sp = "\n- 小标题：4.1、4.2、4.3。必须含至少1个<chart/>和1个<table/>。"
    elif name == "解决方案":
        sp = "\n- 小标题：5.1、5.2、5.3。必须含至少1个<chart/>"
    elif name == "结论":
        sp = "\n- 300-500字总结，无编号。"
    else:
        sp = ""
    cr = build_chart_rule() if name in ["现状分析", "问题与原因分析", "解决方案"] else ""
    hd = "禁止任何标题格式" if name in ["引言", "国内外研究现状", "结论"] else f"小标题：{num}.1 / {num}.2格式"

    revision_section = ""
    if chapter_rev:
        issues = "\n".join([f"- {i}" for i in chapter_rev.get("issues", [])])
        fixes = "\n".join([f"- {f}" for f in chapter_rev.get("fix_requirements", [])])
        keys = "\n".join([f"- {k}" for k in chapter_rev.get("key_points", [])])
        if issues or fixes or keys:
            revision_section = f"""
【重写要求 - 必须严格遵守】
本章在原论文中存在以下问题：
{issues}

修正方向：
{fixes}

必须包含的核心内容：
{keys}
"""

    prompt = f"""写《{profile['title']}》第{num}章 {name}。字数：{MG_WORD_LIMITS[name]}内。
格式铁律：1.{hd} 2.无大标题 3.无markdown符号 4.无加粗 5.纯文本正文{sp}
{cr} 对象：{profile['company']} 行业：{profile['industry']} 只输出正文。{revision_section}"""
    text = call_llm(prompt)
    return re.sub(r'^第[一二三四五六\d]+章.*?\n', '', text).strip()


# ================== [ 重写 TXT 管道 ] ==================
def run_revise_txt_pipeline(task_id: str, profile: dict, revision_plan: dict):
    """修改版论文生成管道：复用原 run_txt_pipeline 结构，融入 revision_plan"""
    folder = f"output/{task_id}"
    os.makedirs(folder, exist_ok=True)

    def update(msg, prog):
        tasks_db[task_id].update({"msg": msg, "progress": prog})
        print(f"[{prog}%] {msg}")

    try:
        update("正在生成摘要...", 5)
        special_notes = revision_plan.get("special_notes", "")
        theory_adjustment = revision_plan.get("theory_adjustment", "")
        abstract_prompt = f"写300字摘要：{profile['title']}，对象{profile['company']}。"
        if special_notes or theory_adjustment:
            abstract_prompt += f"\n\n【特别说明】\n{special_notes}\n{theory_adjustment}"
        abstract = call_llm(abstract_prompt, 600)

        txt = f"{profile['title']}\n\n摘要\n{abstract}\n\n关键词\n{profile['industry']}；{profile['company']}\n\n目录\n\n"

        chapter_revisions = revision_plan.get("chapter_revisions", {})
        for name, num in MG_CHAPTERS:
            chapter_rev = chapter_revisions.get(name)
            txt += f"第{num}章 {name}\n{revise_gen_chapter(name, num, profile, chapter_rev, update)}\n\n"

        update("正在生成参考文献...", 85)
        refs = call_llm(f"围绕{profile['title']}，生成12条规范参考文献[M][J][D]。", 1500)
        txt += "参考文献\n" + "\n".join([r.strip() for r in refs.split('\n') if r.strip().startswith('[')][:12])

        txt_path = os.path.join(folder, "00_完整论文.txt")
        with open(txt_path, 'w', encoding='utf-8') as f:
            f.write(txt)
        update("TXT生成完毕！", 100)
        tasks_db[task_id].update({"status": "completed", "files": [txt_path]})
    except Exception as e:
        tasks_db[task_id].update({"status": "error", "msg": f"生成失败: {str(e)}"})
        print(f"生成失败: {e}")


# ================== [ 主入口 ] ==================
def main():
    parser = argparse.ArgumentParser(description="论文修改程序 - 经管类论文深度重构")
    parser.add_argument("--input", "-i", required=True, help="输入的DOCX文件路径")
    parser.add_argument("--output", "-o", default="./output", help="输出目录（默认./output）")
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"错误：输入文件不存在 {args.input}")
        return

    os.makedirs(args.output, exist_ok=True)

    print("=" * 50)
    print("【论文修改程序】启动")
    print("=" * 50)

    print("\n步骤1/4：提取DOCX文本...")
    original_text = extract_docx_text(args.input)
    if not original_text:
        print("错误：未能从DOCX中提取文本，程序终止。")
        return
    print(f"  提取成功，共 {len(original_text)} 字符")

    print("\n步骤2/4：分析原文画像...")
    analysis_result = analyze_original_paper(original_text)
    print(f"  识别标题：{analysis_result.get('title', 'N/A')}")
    print(f"  识别专业：{analysis_result.get('major', 'N/A')}")
    print(f"  识别对象：{analysis_result.get('company', 'N/A')}")

    print("\n步骤3/4：诊断与重构...")
    revision_plan = diagnose_and_reconstruct(analysis_result)
    report_path = os.path.join(args.output, "诊断报告.txt")
    save_diagnosis_report(revision_plan, report_path)
    print(f"  诊断报告已保存：{report_path}")
    diag = revision_plan.get("diagnosis_report", {})
    print(f"  总体评分：{diag.get('overall_score', 'N/A')}")
    print(f"  共发现 {len(diag.get('issues', []))} 个问题")

    major = revision_plan.get("revised_profile", {}).get("major", analysis_result.get("major", "工商管理"))
    profile = merge_into_standard_profile(revision_plan.get("revised_profile", {}), major)

    task_id = str(uuid.uuid4())
    tasks_db[task_id] = {"status": "running", "progress": 0, "msg": "启动", "files": []}

    print(f"\n步骤4/4：生成修改版论文，任务ID：{task_id}")
    run_revise_txt_pipeline(task_id, profile, revision_plan)

    txt_path = os.path.join("output", task_id, "00_完整论文.txt")
    if not os.path.exists(txt_path):
        print("错误：TXT生成失败，请检查日志。")
        return

    docx_path = os.path.join("output", task_id, "论文_修改版.docx")

    def update_fn(msg, prog):
        print(f"[DOCX {prog}%] {msg}")

    txt_to_docx_safe(txt_path, docx_path, update_fn)

    final_docx = os.path.join(args.output, "论文_修改版.docx")
    shutil.move(docx_path, final_docx)
    print(f"\n{'=' * 50}")
    print(f"全部完成！")
    print(f"诊断报告：{report_path}")
    print(f"修改版论文：{final_docx}")
    print(f"{'=' * 50}")


if __name__ == "__main__":
    main()

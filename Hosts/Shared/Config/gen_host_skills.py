# -*- coding: utf-8 -*-
"""
生成各宿主专属技能子集 (word_skills.json / ppt_skills.json)。

读取内置 Skills/*.md（YAML frontmatter + 正文），按宿主能力筛选，
输出与 SkillBridge.GetAllSkills 兼容的技能 JSON（camelCase 字段）。
宿主工程不引用主 Excel 程序集，故用预生成静态文件由 HostBridgeBase 直接读取。

用法: python gen_host_skills.py
依赖: 仅需标准库 (json)。
"""
import json
import os
import re

# frontmatter 元数据的取值类型
MetaValue = str | bool | list[str]

ROOT = os.path.dirname(os.path.abspath(__file__))
SKILLS_DIR = os.path.normpath(os.path.join(ROOT, "..", "..", "..", "Skills"))
# 注意：本脚本放在 Hosts/Shared/Config/ 下，仓库根为 ../../../
SKIP_DIRS = ("Engine", "Testing")

# 生成哪些宿主的技能子集（Excel 走主程序集 SkillBridge，不在此生成）。
# 每个技能的适用宿主由其 .md frontmatter 的 hosts: 字段声明。
TARGET_HOSTS: list[str] = ["word", "ppt"]

# 标记为"快捷入口"（欢迎栏展示）；若技能 frontmatter 自带 isQuickShortcut 则以之优先
QUICK_SHORTCUTS: set[str] = {"SmartSummary", "SmartDataAnalyzer"}

# frontmatter 别名归一化（frontmatter 用 powerpoint，输出用 ppt）
HOST_ALIASES: dict[str, str] = {"powerpoint": "ppt"}


def parse_frontmatter(text: str) -> tuple[dict[str, MetaValue], str]:
    """解析简单 YAML frontmatter，返回 (meta, body)。"""
    lines: list[str] = text.split("\n")
    if not lines or lines[0].strip() != "---":
        return {}, text
    end: int | None = None
    for i in range(1, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
    if end is None:
        return {}, text
    fm_lines: list[str] = lines[1:end]
    body: str = "\n".join(lines[end + 1:])
    meta: dict[str, MetaValue] = {}
    cur_key: str | None = None
    for ln in fm_lines:
        stripped: str = ln.strip()
        if cur_key is not None and stripped.startswith("- ") and isinstance(meta.get(cur_key), list):
            value = meta[cur_key]
            if isinstance(value, list):
                value.append(stripped[2:].strip())
            continue
        if ":" in ln and not stripped.startswith("- "):
            k: str
            v: str
            k, _, v = ln.partition(":")
            k = k.strip()
            v = v.strip()
            if v == "":
                cur_key = k
                meta[k] = []
            else:
                cur_key = k
                if k == "hosts":
                    # 解析内联列表 hosts: [excel, word, powerpoint]
                    inner: str = v.strip().strip("[").strip("]").strip()
                    meta[k] = [x.strip() for x in inner.split(",") if x.strip()] if inner else []
                elif v.lower() == "true":
                    meta[k] = True
                elif v.lower() == "false":
                    meta[k] = False
                else:
                    meta[k] = v
    return meta, body.strip()


def _has_excelbridge(deps) -> bool:
    if deps is None:
        return False
    if isinstance(deps, str):
        return "excelbridge" in deps.lower()
    if isinstance(deps, list):
        return any("excelbridge" in str(d).lower() for d in deps)
    return False


def _is_excel_only(meta: dict[str, MetaValue]) -> bool:
    """Excel-only 强约束判定（防非 Excel 宿主卡死）。

    仅当技能「spec 引擎渲染 Excel VBA」时才强制 Excel-only，需同时满足：
      (1) executionMode=spec 且依赖 ExcelBridge；
      (2) name 不属于「宿主生成类」（生成 PPT/Word 的 spec 技能在自身宿主内安全，
          如 GeneratePPTReport 渲染 PPT，允许进入 ppt 桶）。
    反例：SmartTableBeautifier 之类 Excel 表操作 spec 技能，一旦误声明 hosts 含
    word/ppt，会被 spec 引擎渲染 Excel VBA 在 Word/PPT 中执行 → 宿主卡死。
    """
    name = str(meta.get("name", ""))
    # 宿主生成类（生成 PPT/Word 的 spec 技能）在自身宿主内安全，不强制 Excel-only
    if re.search(r'ppt|word|幻灯片|文档|汇报', name, re.I):
        return False
    em = meta.get("executionMode")
    if isinstance(em, str) and em.strip().lower() == "spec" and _has_excelbridge(meta.get("dependencies")):
        return True
    return False


def build_skill(
    meta: dict[str, MetaValue], body: str, fallback_name: str = ""
) -> dict[str, object] | None:
    name: str = str(meta.get("name") or fallback_name)
    if not name:
        return None
    display_name: str = str(meta.get("displayName") or "")
    if not display_name:
        # 无 frontmatter 的技能（如 ThinkingFramework）：用首个 # 标题
        for ln in body.split("\n"):
            s: str = ln.strip()
            if s.startswith("#"):
                display_name = s.lstrip("#").strip()
                break
    if not display_name:
        display_name = name
    qs_meta = meta.get("isQuickShortcut")
    is_qs: bool = bool(qs_meta) if isinstance(qs_meta, bool) else (name in QUICK_SHORTCUTS)
    return {
        "name": name,
        "displayName": display_name,
        "description": meta.get("description", ""),
        "keywords": meta.get("keywords", []) if isinstance(meta.get("keywords"), list) else [],
        "examples": meta.get("examples", []) if isinstance(meta.get("examples"), list) else [],
        "category": meta.get("category", "其他"),
        "isGeneric": bool(meta.get("isGeneric", False)),
        "needsConfig": bool(meta.get("needsConfig", False)),
        "pipeline": meta.get("pipeline", ""),
        "executionMode": meta.get("executionMode", ""),
        "processingMode": "Skill",
        "isCustom": False,
        "isEnabled": True,
        "isQuickShortcut": is_qs,
        "version": 1,
        "semanticVersion": "1.0.0",
        "allowedTools": [],
        "promptContent": body,
        # agent 模式（executionMode=agent）会把 llmSystemPrompt 作为 LLM 系统提示注入 agent-loop；
        # 若不单独写出，前端只会回退到 description，正文里的详细调用指南便失效。
        "llmSystemPrompt": body,
    }


def main() -> None:
    buckets: dict[str, list[dict[str, object]]] = {h: [] for h in TARGET_HOSTS}
    for fn in sorted(os.listdir(SKILLS_DIR)):
        if not fn.endswith(".md") or fn.startswith("__") or fn in SKIP_DIRS:
            continue
        p: str = os.path.join(SKILLS_DIR, fn)
        with open(p, "r", encoding="utf-8") as f:
            text: str = f.read()
        meta, body = parse_frontmatter(text)
        name: str = str(meta.get("name") or fn[:-3])
        # 归一化：frontmatter 用 powerpoint，输出用 ppt
        raw_hosts = meta.get("hosts")
        hosts: list[str] = (
            [str(HOST_ALIASES.get(x, x)) for x in raw_hosts]
            if isinstance(raw_hosts, list)
            else []
        )
        for h in TARGET_HOSTS:
            if h in hosts:
                # Excel-only 强约束：SkillSpec/ExcelBridge 技能严禁进入 Word/PPT 桶
                if _is_excel_only(meta):
                    print("[WARN] 跳过 Excel-only 技能 %s 进入 %s 桶"
                          "（executionMode=spec 或依赖 ExcelBridge，会致非 Excel 宿主卡死）" % (name, h))
                    continue
                sk = build_skill(meta, body, fallback_name=name)
                if sk is not None:
                    buckets[h].append(sk)
    for h in TARGET_HOSTS:
        out: str = os.path.join(ROOT, h + "_skills.json")
        with open(out, "w", encoding="utf-8") as f:
            json.dump(buckets[h], f, ensure_ascii=False, indent=2)
        print("[OK] %s -> %s (%d 个技能)" % (h, out, len(buckets[h])))


if __name__ == "__main__":
    main()

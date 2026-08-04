import json
import os
import re
from pathlib import Path
from datetime import datetime

SESSIONS_DIR = Path(os.path.expandvars(r"%USERPROFILE%\.codex\sessions"))
ARCHIVED_DIR = Path(os.path.expandvars(r"%USERPROFILE%\.codex\archived_sessions"))
NOTES_DIR = Path(r"C:\Users\pc\ZCodeProject\knowledge-base\notes")
NOTES_DIR.mkdir(parents=True, exist_ok=True)

def slugify(text):
    """生成文件名 slug"""
    text = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]+', '-', text).strip('-')
    return text[:80] or "untitled"

def extract_conversation(filepath):
    """从 JSONL 提取对话内容"""
    messages = []
    session_id = ""
    session_date = ""
    cwd = ""
    
    with open(filepath, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                record = json.loads(line)
            except json.JSONDecodeError:
                continue
            
            payload = record.get("payload", {})
            rtype = record.get("type", "")
            
            # 提取元信息
            if rtype == "session_meta":
                session_id = payload.get("session_id", "")
                session_date = payload.get("timestamp", "")
                cwd = payload.get("cwd", "")
            
            # 提取用户消息
            if rtype == "event_msg" and payload.get("type") == "user_message":
                msg = payload.get("message", "").strip()
                if msg:
                    messages.append({"role": "user", "content": msg})
            
            # 提取 agent 回复
            if rtype == "event_msg" and payload.get("type") == "agent_message":
                msg = payload.get("message", "").strip()
                if msg:
                    messages.append({"role": "assistant", "content": msg})
            
            # 也检查 response_item 中的消息
            if rtype == "response_item":
                item = payload
                if item.get("type") == "message" and item.get("role") == "user":
                    for c in item.get("content", []):
                        if c.get("type") == "input_text":
                            txt = c.get("text", "").strip()
                            if txt:
                                messages.append({"role": "user", "content": txt})
                if item.get("type") == "message" and item.get("role") == "assistant":
                    for c in item.get("content", []):
                        if c.get("type") == "output_text":
                            txt = c.get("text", "").strip()
                            if txt:
                                messages.append({"role": "assistant", "content": txt})
    
    return {
        "session_id": session_id,
        "date": session_date,
        "cwd": cwd,
        "messages": messages
    }

def format_markdown(conv):
    """生成 Markdown 笔记"""
    # 标题：用第一条用户消息
    title = "未命名对话"
    skip_keywords = ["<permissions", "<environment_context>", "<app-context>", "<skills_instructions>",
                     "<collaboration_mode>", "Filesystem sandboxing"]
    for m in conv["messages"]:
        if m["role"] == "user":
            first_line = m["content"].split("\n")[0][:80]
            # 跳过系统模板消息
            if any(first_line.startswith(kw) for kw in skip_keywords):
                continue
            title = first_line[:60]
            break
    
    date_str = conv["date"][:10] if conv["date"] else "未知日期"
    sid = conv["session_id"][:8] if conv["session_id"] else ""
    
    lines = [f"# {title}", ""]
    lines.append(f"- **日期**: {date_str}")
    lines.append(f"- **会话ID**: `{sid}`")
    if conv["cwd"]:
        lines.append(f"- **工作目录**: `{conv['cwd']}`")
    lines.append("")
    lines.append("---")
    lines.append("")
    
    for m in conv["messages"]:
        role_label = "🧑 用户" if m["role"] == "user" else "🤖 Codex"
        content = m["content"]
        # 截断过长的消息
        if len(content) > 3000:
            content = content[:3000] + "\n\n... *(内容过长已截断)*"
        lines.append(f"### {role_label}")
        lines.append("")
        lines.append(content)
        lines.append("")
    
    return "\n".join(lines), title

def main():
    # 收集所有会话（去重：活跃 > 归档）
    seen = set()
    files_to_import = []
    
    # 先加活跃会话
    for f in sorted(SESSIONS_DIR.rglob("*.jsonl")):
        if "2026" in str(f):
            sid = f.stem.split("-")[-1] if "-" in f.stem else f.stem
            if sid not in seen:
                seen.add(sid)
                files_to_import.append(f)
    
    # 再加归档（跳过已存在的）
    for f in sorted(ARCHIVED_DIR.glob("*.jsonl")):
        sid = f.stem.split("-")[-1] if "-" in f.stem else f.stem
        if sid not in seen:
            seen.add(sid)
            files_to_import.append(f)
    
    imported = 0
    skipped = 0
    
    for filepath in files_to_import:
        try:
            conv = extract_conversation(filepath)
            if not conv["messages"]:
                skipped += 1
                continue
            
            md, title = format_markdown(conv)
            slug = slugify(title)
            note_path = NOTES_DIR / f"{slug}.md"
            
            # 处理重名
            counter = 2
            original_slug = slug
            while note_path.exists():
                slug = f"{original_slug}-{counter}"
                note_path = NOTES_DIR / f"{slug}.md"
                counter += 1
            
            note_path.write_text(md, encoding="utf-8")
            imported += 1
            print(f"✅ {title[:50]} → {note_path.name}")
        except Exception as e:
            print(f"❌ {filepath.name}: {e}")
    
    print(f"\n完成: {imported} 篇导入, {skipped} 篇跳过")

if __name__ == "__main__":
    main()


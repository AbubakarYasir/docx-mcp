"""DOCX_MCP_AUTHOR sets the name Word shows for tracked changes and comments."""

from __future__ import annotations

import os
import subprocess
import sys


def test_env_sets_default_author(tmp_path):
    code = (
        "from docx_mcp.document import DocxDocument\n"
        "W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'\n"
        "W14='{http://schemas.microsoft.com/office/word/2010/wordml}'\n"
        f"d=DocxDocument.create(r'{tmp_path / 'a.docx'}')\n"
        "p=d._trees['word/document.xml'].find(f'.//{W}p').get(f'{W14}paraId')\n"
        "q=d.insert_paragraph(p,'Hello')['para_id']\n"
        "d.insert_text(q,' world')\n"
        "d.add_comment(q,'note')\n"
        "ins=d._trees['word/document.xml'].find(f'.//{W}ins')\n"
        "com=d._trees['word/comments.xml'].find(f'{W}comment')\n"
        "print(ins.get(f'{W}author')); print(com.get(f'{W}author'))\n"
    )
    env = {**os.environ, "DOCX_MCP_AUTHOR": "Abu Bakar Gondal"}
    out = subprocess.run(
        [sys.executable, "-c", code], env=env, capture_output=True, text=True, check=True
    ).stdout.split("\n")
    assert out[0] == "Abu Bakar Gondal"
    assert out[1] == "Abu Bakar Gondal"


def test_default_is_claude_without_env(tmp_path):
    env = {k: v for k, v in os.environ.items() if k != "DOCX_MCP_AUTHOR"}
    out = subprocess.run(
        [
            sys.executable,
            "-c",
            "from docx_mcp.document.base import DEFAULT_AUTHOR;print(DEFAULT_AUTHOR)",
        ],
        env=env,
        capture_output=True,
        text=True,
        check=True,
    ).stdout.strip()
    assert out == "Claude"

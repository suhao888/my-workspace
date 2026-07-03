"""Test formula and metadata specifically."""

import subprocess
import json

proc = subprocess.Popen(
    ["python", "-m", "excel_mcp", "stdio"],
    stdin=subprocess.PIPE,
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
)

batch = ""
batch += (
    json.dumps(
        {
            "jsonrpc": "2.0",
            "id": 1,
            "method": "initialize",
            "params": {
                "protocolVersion": "2024-11-05",
                "capabilities": {},
                "clientInfo": {"name": "t", "version": "1"},
            },
        }
    )
    + "\n"
)

# Read after write
batch += (
    json.dumps(
        {
            "jsonrpc": "2.0",
            "id": 2,
            "method": "tools/call",
            "params": {
                "name": "read_data_from_excel",
                "arguments": {
                    "filepath": "E:\\Projects\\my-workspace\\test_mcp.xlsx",
                    "sheet_name": "Sheet1",
                    "range": "A1:C5",
                },
            },
        }
    )
    + "\n"
)

# Apply formula
batch += (
    json.dumps(
        {
            "jsonrpc": "2.0",
            "id": 3,
            "method": "tools/call",
            "params": {
                "name": "apply_formula",
                "arguments": {
                    "filepath": "E:\\Projects\\my-workspace\\test_mcp.xlsx",
                    "sheet_name": "Sheet1",
                    "cell": "C4",
                    "formula": "=C2-C3",
                },
            },
        }
    )
    + "\n"
)

# Metadata
batch += (
    json.dumps(
        {
            "jsonrpc": "2.0",
            "id": 4,
            "method": "tools/call",
            "params": {
                "name": "get_workbook_metadata",
                "arguments": {"filepath": "E:\\Projects\\my-workspace\\test_mcp.xlsx"},
            },
        }
    )
    + "\n"
)

stdout, stderr = proc.communicate(input=batch.encode(), timeout=10)

for line in stdout.decode().strip().split("\n"):
    if not line.strip():
        continue
    try:
        obj = json.loads(line)
        rid = obj.get("id")
        if "error" in obj:
            err = obj["error"]
            msg = err.get("message", str(err))[:300]
            print(f"[{rid}] ERROR: {msg}")
        elif "result" in obj:
            content = obj["result"].get("content", [{}])
            text = (
                content[0].get("text", "")[:300]
                if content
                else str(obj["result"])[:300]
            )
            print(f"[{rid}] OK: {text}")
    except Exception as e:
        print(f"PARSE: {line[:100]} -> {e}")

"""测试 PPStructure 表格识别"""

import os

os.environ["FLAGS_use_mkldnn"] = "0"
import torch

from paddleocr import PPStructure

engine = PPStructure(show_log=False, lang="ch")
print("PPStructure 引擎就绪")

result = engine("E:/Projects/my-workspace/pic2table/input/test_table.png")
print(f"结果数量: {len(result)}")

for i, block in enumerate(result):
    print(f"\n--- Block {i} ---")
    btype = block.get("type", "?")
    print(f"  类型: {btype}")
    if btype == "table":
        html = block.get("res", {}).get("html", "")
        print(f"  HTML (前300字): {html[:300]}")
    elif btype == "text":
        print(f"  内容: {block.get('res', '')[:200]}")
    else:
        print(f"  完整block: {str(block)[:300]}")

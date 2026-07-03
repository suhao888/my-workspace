"""测试 PPStructure - 使用 server 版模型"""

import os

os.environ["FLAGS_use_mkldnn"] = "0"
import torch

from paddleocr import PPStructure

# 使用 server 版模型（更精确）
engine = PPStructure(
    show_log=False,
    lang="ch",
    # 使用 server 版表格结构模型
    table_model_dir=None,  # 自动下载 server 版
    # 调大识别区域
    det_db_thresh=0.3,
    det_db_box_thresh=0.5,
    # 表格参数
    table_max_len=960,
)
print("PPStructure server 引擎就绪")

result = engine("E:/Projects/my-workspace/pic2table/input/test_table.png")
print(f"结果数量: {len(result)}")

for i, block in enumerate(result):
    btype = block.get("type", "?")
    print(f"\n--- Block {i} --- 类型: {btype}")
    if btype == "table":
        html = block.get("res", {}).get("html", "")
        print(f"  HTML: {html[:500]}")
    elif btype == "text":
        print(f"  内容: {block.get('res', '')[:200]}")

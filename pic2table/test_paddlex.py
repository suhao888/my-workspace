"""测试 PaddleX 表格识别 - 独立脚本，设置好环境变量再导入"""

import os
import sys

# 在导入任何 Paddle 相关库之前设置环境变量
os.environ["FLAGS_use_mkldnn"] = "0"
os.environ["DNNL_DEFAULT_FPMATH_MODE"] = "BF16"

# 验证环境变量生效
print(f"FLAGS_use_mkldnn = {os.environ.get('FLAGS_use_mkldnn', 'NOT SET')}")

from paddlex import create_pipeline

print("创建 table_recognition 管线...")
pipeline = create_pipeline("table_recognition")

print("预测图片...")
result = pipeline.predict("E:/Projects/my-workspace/pic2table/input/test_table.png")

for res in result:
    print(f"\n结果类型: {type(res).__name__}")
    for attr in sorted(dir(res)):
        if not attr.startswith("_") and attr not in ("contour",):
            try:
                val = getattr(res, attr)
                if val is not None:
                    s = str(val)
                    print(f"  {attr}: {s[:400]}")
            except Exception as e:
                print(f"  {attr}: [ERR {e}]")

print("\n完成")

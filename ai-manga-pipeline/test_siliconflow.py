"""测试 SiliconFlow 图片生成 - 找能用的模型"""

import os
import requests
import base64
from io import BytesIO
from PIL import Image

env_path = r"E:\Projects\my-workspace\ai-manga-pipeline\.env"
api_key = None
with open(env_path, "r", encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if line.startswith("SILICONFLOW_API_KEY="):
            api_key = line.split("=", 1)[1].strip().strip('"').strip("'")
            break

headers = {"Authorization": f"Bearer {api_key}"}
prompt = "古代女子穿淡蓝色长裙在花园中，唯美风格"
base_url = "https://api.siliconflow.cn/v1"

# 测试多个模型
models_to_test = [
    "Kwai-Kolors/Kolors",
    "Qwen/Qwen-Image",
    "Tongyi-MAI/Z-Image-Turbo",
    "Tongyi-MAI/Z-Image",
    "baidu/ERNIE-Image-Turbo",
]

for model in models_to_test:
    print(f"\n测试模型: {model}")
    # 尝试 OpenAI 兼容格式
    payload = {
        "model": model,
        "prompt": prompt,
        "width": 1024,
        "height": 1024,
        "n": 1,
        "response_format": "b64_json",
    }

    try:
        resp = requests.post(
            f"{base_url}/images/generations", headers=headers, json=payload, timeout=120
        )
        print(f"  OpenAI 格式: {resp.status_code}")

        if resp.status_code == 200:
            data = resp.json()
            if "data" in data and len(data["data"]) > 0:
                img_data = data["data"][0]
                if "b64_json" in img_data:
                    img = Image.open(BytesIO(base64.b64decode(img_data["b64_json"])))
                    out = os.path.join(
                        r"E:\Projects\my-workspace\ai-manga-pipeline\output",
                        f"test_{model.split('/')[-1]}.png",
                    )
                    img.save(out)
                    print(f"  ✅ 生成成功！保存到: {out}")
                    print(f"  尺寸: {img.size}")
                    break
                elif "url" in img_data:
                    print(f"  返回 URL: {img_data['url'][:80]}")
                    break
            else:
                print(f"  响应: {str(data)[:200]}")
        elif resp.status_code == 400:
            err = resp.json()
            print(f"  错误: {err.get('message', str(err)[:100])}")
        else:
            print(f"  响应: {resp.text[:200]}")
    except Exception as e:
        print(f"  异常: {e}")
else:
    print("\n❌ 所有模型都失败了，尝试其他 endpoint...")

    # 尝试 SiliconFlow 原生格式
    for model in models_to_test[:2]:
        print(f"\n尝试原生格式 - {model}")
        payload = {
            "model": model,
            "prompt": prompt,
            "width": 1024,
            "height": 1024,
            "seed": 42,
        }
        try:
            resp = requests.post(
                f"{base_url}/images/generations",
                headers=headers,
                json=payload,
                timeout=120,
            )
            print(f"  状态码: {resp.status_code}, 响应: {resp.text[:200]}")
        except Exception as e:
            print(f"  异常: {e}")

"""测试 SiliconFlow Wan2.2 视频 API"""

import os, requests, time, base64

env_path = r"E:\Projects\my-workspace\ai-manga-pipeline\.env"
api_key = None
with open(env_path, "r", encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if line.startswith("SILICONFLOW_API_KEY="):
            api_key = line.split("=", 1)[1].strip().strip('"').strip("'")
            break

headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

# 找一个测试图片
test_img = (
    r"E:\Projects\my-workspace\ai-manga-pipeline\output\temp\images\scene_001.png"
)
if not os.path.exists(test_img):
    print(f"测试图片不存在: {test_img}")
    # 生成一个临时测试图
    from PIL import Image, ImageDraw

    img = Image.new("RGB", (720, 1280), (50, 100, 150))
    img.save(test_img)
    print("已生成临时测试图")

with open(test_img, "rb") as f:
    img_b64 = base64.b64encode(f.read()).decode()

# 提交视频任务
print("提交 Wan2.2 图生视频任务...")
payload = {
    "model": "Wan-AI/Wan2.2-I2V-A14B",
    "prompt": "人物轻微动作，画面自然流畅",
    "image_size": "720x1280",
    "image": f"data:image/png;base64,{img_b64}",
}

resp = requests.post(
    "https://api.siliconflow.cn/v1/video/submit",
    headers=headers,
    json=payload,
    timeout=30,
)

print(f"提交状态: {resp.status_code}")
if resp.status_code != 200:
    print(f"错误: {resp.text[:300]}")
    exit()

data = resp.json()
request_id = data.get("requestId")
print(f"requestId: {request_id}")

# 轮询状态
print("\n轮询生成状态...")
for i in range(60):
    time.sleep(5)
    status_resp = requests.post(
        "https://api.siliconflow.cn/v1/video/status",
        headers=headers,
        json={"requestId": request_id},
        timeout=15,
    )
    if status_resp.status_code != 200:
        print(f"  [{i * 5}s] 状态查询失败: {status_resp.status_code}")
        continue

    result = status_resp.json()
    status = result.get("status")
    print(f"  [{i * 5}s] status={status}")

    if status == "Succeed":
        videos = result.get("results", {}).get("videos", [])
        if videos:
            url = videos[0]["url"]
            print(f"\n✅ 视频生成成功!")
            print(f"URL: {url}")
            # 下载
            out = r"E:\Projects\my-workspace\ai-manga-pipeline\output\test_wan2_2.mp4"
            dl = requests.get(url, timeout=60)
            with open(out, "wb") as f:
                f.write(dl.content)
            print(f"已下载到: {out} ({len(dl.content) / 1024:.0f} KB)")
        break
    elif status == "Failed":
        print(f"\n❌ 失败: {result.get('reason', '未知')}")
        break
    elif status in ("InQueue", "InProgress"):
        continue
    else:
        print(f"  未知状态: {result}")
else:
    print(f"\n超时")

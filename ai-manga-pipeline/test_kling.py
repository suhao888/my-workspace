"""测试可灵 Kling - 逐个排查"""

import os, requests, json, base64, time

env_path = r"E:\Projects\my-workspace\ai-manga-pipeline\.env"
api_key = None
with open(env_path, "r", encoding="utf-8") as f:
    for line in f:
        line = line.strip()
        if line.startswith("KLING_API_KEY="):
            api_key = line.split("=", 1)[1].strip().strip('"').strip("'")
            break

# 从 key 判断格式
print(f"API Key: {api_key[:25]}...")
print(f"格式: ", end="")
if "|" in api_key:
    print("AK|SK 格式")
    ak, sk = api_key.split("|", 1)
elif api_key.startswith("api-key-kling-"):
    print("Bearer 格式")
    ak, sk = api_key, api_key
else:
    print("未知格式")

# 测试 1: 纯 GET 请求看看 key 是否有效
headers = {"Authorization": f"Bearer {api_key}"}
resp = requests.get("https://api.klingai.com/v1/user", headers=headers, timeout=15)
print(f"\nGET /v1/user: {resp.status_code} {resp.text[:200]}")

# 测试 2: 可能 key 是 AK 和 SK 组合
# 用整个 key 同时作为 AK 和 SK
import hashlib, hmac


def sign_v2(ak, sk, method, path, body=""):
    now = str(int(time.time()))
    host = "api.klingai.com"

    # 待签名的 headers
    signed_headers = {
        "host": host,
        "x-api-timestamp": now,
        "Content-Type": "application/json",
    }

    # 构建 header string
    hdr_keys = sorted(signed_headers.keys())
    header_str = "".join(f"{k}:{signed_headers[k]}\n" for k in hdr_keys)

    # 构建签名串
    sign_str = f"{method}\n{path}\n\n{header_str}\n{body}"

    signature = hmac.new(sk.encode(), sign_str.encode(), hashlib.sha256).hexdigest()

    return {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {ak}",
        "x-api-timestamp": now,
        "x-api-signature": signature,
    }


# 尝试 v2 签名
for endpoint in ["/v1/videos/image2video", "/v1/images/image2video"]:
    body = json.dumps(
        {
            "model": "kling-v1.6",
            "image_url": "data:image/png;base64,dGVzdA==",
            "prompt": "action",
        }
    )
    hdrs = sign_v2(ak, sk, "POST", endpoint, body)
    resp = requests.post(
        f"https://api.klingai.com{endpoint}", headers=hdrs, data=body, timeout=15
    )
    print(f"\nPOST {endpoint} (v2签名): {resp.status_code} {resp.text[:200]}")

# 测试 3: 直接用 Bearer 但换个 endpoint
for ep in ["/v1/images/generations", "/v1/video/generate", "/v1/text2video"]:
    payload = {"model": "kling-v1.6", "prompt": "test"}
    resp = requests.post(
        f"https://api.klingai.com{ep}", headers=headers, json=payload, timeout=15
    )
    print(f"POST {ep}: {resp.status_code} {resp.text[:100]}")

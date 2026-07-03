"""
剧本生成模块 — 用 DeepSeek 生成结构化剧本
"""

import json
import re
from typing import List, Optional
from dataclasses import dataclass, field, asdict
from openai import OpenAI

from config import config


@dataclass
class CharacterDef:
    name: str
    appearance: str
    personality: str


@dataclass
class DialogueLine:
    speaker: str  # 角色名 或 "旁白"
    text: str


@dataclass
class SceneDef:
    id: int
    location: str
    characters: List[str]  # 本场出现的角色名列表
    scene_description: str  # 场景描述（给剪辑看）
    image_prompt: str  # 画面描述（给 AI 出图用，中文）
    dialogue: List[DialogueLine] = field(default_factory=list)
    duration: float = 5.0
    # 运行时填充的字段
    image_path: Optional[str] = None
    video_path: Optional[str] = None


@dataclass
class Script:
    title: str
    genre: str
    characters: List[CharacterDef]
    scenes: List[SceneDef]


SYSTEM_PROMPT = """你是一个AI漫剧编剧。你的任务是写"AI漫剧"的分镜剧本。

AI漫剧的特点：
- 用 AI 生成静态画面 + 配音 + 运镜特效，做成剧情短视频
- 每集 2-3 分钟，约 10-15 个分镜
- 题材以网文改编为主：穿越、重生、霸总、神医、末世、兽世、娱乐圈等
- 节奏快，每镜 5-10 秒，对话密集
- 画面描述要具体，便于 AI 出图

输出格式要求：
你必须输出合法的 JSON，不要包含代码块标记或其他文字：
{
  "title": "剧名",
  "genre": "题材",
  "characters": [
    {"name": "角色名", "appearance": "外貌描述", "personality": "性格描述"}
  ],
  "scenes": [
    {
      "id": 1,
      "location": "场景位置",
      "characters": ["出场角色名"],
      "scene_description": "场景描述（给剪辑参考）",
      "image_prompt": "画面描述（30-50字，包含角色外貌、场景、氛围，用于AI出图）",
      "dialogue": [
        {"speaker": "角色名或旁白", "text": "台词"}
      ],
      "duration": 5.0
    }
  ]
}

要求：
1. 角色外貌描述必须在 image_prompt 中重复出现，确保 AI 出图一致性
2. image_prompt 用中文，包含：角色外貌、动作、场景、色调/氛围
3. 每个分镜至少 1 句对话或旁白
4. 开头要有吸引人的钩子（hook），结尾要有悬念
5. 总时长控制在 120-180 秒（约 12-18 个分镜）"""


class ScriptGenerator:
    """用 DeepSeek 生成剧本"""

    def __init__(self):
        self.client = OpenAI(api_key=config.llm.api_key, base_url=config.llm.base_url)
        self.model = config.llm.model

    def generate(self, genre: str, premise: str, num_scenes: int = 15) -> Script:
        """根据题材和梗概生成完整剧本"""
        user_prompt = f"""题材：{genre}
故事梗概：{premise}

请生成一个完整的 AI 漫剧剧本，包含 {num_scenes} 个分镜。
每个分镜的 image_prompt 要详细描述画面，包含角色外貌、场景、色调。"""

        resp = self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=config.llm.temperature,
            max_tokens=config.llm.max_tokens,
            response_format={"type": "json_object"},
        )

        raw = resp.choices[0].message.content
        return self._parse(raw)

    def _parse(self, raw: str) -> Script:
        """解析 JSON 响应为 Script 对象"""
        raw = raw.strip()
        # 清理可能的 markdown 代码块
        raw = re.sub(r"^```(?:json)?\s*", "", raw)
        raw = re.sub(r"\s*```$", "", raw)
        data = json.loads(raw)
        chars = [CharacterDef(**c) for c in data.get("characters", [])]
        scenes = []
        for s in data.get("scenes", []):
            dialogues = [DialogueLine(**d) for d in s.get("dialogue", [])]
            scenes.append(
                SceneDef(
                    id=s["id"],
                    location=s.get("location", ""),
                    characters=s.get("characters", []),
                    scene_description=s.get("scene_description", ""),
                    image_prompt=s.get("image_prompt", ""),
                    dialogue=dialogues,
                    duration=s.get("duration", 5.0),
                )
            )
        return Script(
            title=data.get("title", "未命名"),
            genre=data.get("genre", ""),
            characters=chars,
            scenes=scenes,
        )

    def save(self, script: Script, path: str):
        """保存剧本到 JSON 文件"""
        data = {
            "title": script.title,
            "genre": script.genre,
            "characters": [asdict(c) for c in script.characters],
            "scenes": [
                {**asdict(s), "dialogue": [asdict(d) for d in s.dialogue]}
                for s in script.scenes
            ],
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def load(self, path: str) -> Script:
        """从 JSON 文件加载剧本"""
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        chars = [CharacterDef(**c) for c in data.get("characters", [])]
        scenes = []
        for s in data.get("scenes", []):
            dialogues = [DialogueLine(**d) for d in s.get("dialogue", [])]
            scenes.append(
                SceneDef(
                    id=s["id"],
                    location=s.get("location", ""),
                    characters=s.get("characters", []),
                    scene_description=s.get("scene_description", ""),
                    image_prompt=s.get("image_prompt", ""),
                    dialogue=dialogues,
                    duration=s.get("duration", 5.0),
                )
            )
        return Script(
            title=data.get("title", "未命名"),
            genre=data.get("genre", ""),
            characters=chars,
            scenes=scenes,
        )

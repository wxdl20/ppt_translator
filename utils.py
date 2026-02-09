import os
from typing import Optional, Tuple

import google.generativeai as genai
from openai import OpenAI


def build_system_prompt(
    source_lang: str,
    target_lang: str,
    topic_hint: Optional[str] = None,
) -> str:
    base_prompt = (
        "You are a theological and academic translator. "
        f"Translate the following text from {source_lang} to {target_lang}. "
        "Preserve the tone. Do not output markdown, just the raw translated text."
    )
    if topic_hint:
        return f"{base_prompt} Document topic: {topic_hint}"
    return base_prompt


def create_client(api_key: str, base_url: str) -> OpenAI:
    return OpenAI(api_key=api_key, base_url=base_url)


def create_gemini_model(api_key: str, model_name: str, system_prompt: str):
    genai.configure(api_key=api_key)
    return genai.GenerativeModel(model_name=model_name, system_instruction=system_prompt)


def translate_text(client: OpenAI, model_name: str, system_prompt: str, text: str) -> str:
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text},
        ],
        temperature=0.3,
    )
    content = response.choices[0].message.content if response.choices else None
    if not content:
        raise ValueError("Empty response from model")
    return content.strip()


def translate_text_gemini(model, text: str) -> str:
    response = model.generate_content(text)
    if not response or not response.text:
        raise ValueError("Empty response from model")
    return response.text.strip()


def load_laozhang_config() -> Tuple[Optional[str], str]:
    _load_env_file()
    api_key = os.getenv("LAOZHANG_API_KEY")
    base_url = os.getenv("LAOZHANG_BASE_URL", "https://api.laozhang.ai/v1")
    return api_key, base_url


def load_gemini_key() -> Optional[str]:
    _load_env_file()
    return os.getenv("GEMINI_API_KEY")


def _load_env_file(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value

from io import BytesIO

import streamlit as st

from translator_engine import DocxStrategy, PptxStrategy, TranslatorEngine
from utils import (
    build_system_prompt,
    create_client,
    create_gemini_model,
    load_gemini_key,
    load_laozhang_config,
    translate_text,
    translate_text_gemini,
)


LANG_OPTIONS = [
    "Chinese",
    "English",
    "Indonesian",
    "Armenian",
    "Nepali",
]

PROVIDER_OPTIONS = [
    "LaoZhang(OpenAI兼容)",
    "Gemini(官方)",
]

LAOZHANG_MODELS = [
    "gemini-2.5-flash-preview",
]

GEMINI_MODELS = [
    "gemini-2.5-flash-preview",
]


def _progress_handler(progress_bar):
    def _inner(done: int, total_steps: int) -> None:
        if total_steps <= 0:
            return
        progress_bar.progress(min(done / total_steps, 1.0))

    return _inner


def _build_engine() -> TranslatorEngine:
    return TranslatorEngine(
        strategies={
            "docx": DocxStrategy(),
            "pptx": PptxStrategy(),
        }
    )


def _translate_file(
    filename: str,
    file_bytes: bytes,
    source_lang: str,
    target_lang: str,
    topic_hint: str,
    provider: str,
    model_name: str,
    laozhang_api_key: str,
    base_url: str,
    gemini_api_key: str,
    progress_bar,
) -> BytesIO:
    system_prompt = build_system_prompt(source_lang, target_lang, topic_hint)
    if provider == "Gemini(官方)":
        model = create_gemini_model(
            api_key=gemini_api_key,
            model_name=model_name,
            system_prompt=system_prompt,
        )

        def translate_one(text: str) -> str:
            return translate_text_gemini(model, text)
    else:
        client = create_client(api_key=laozhang_api_key, base_url=base_url)

        def translate_one(text: str) -> str:
            return translate_text(client, model_name, system_prompt, text)

    engine = _build_engine()
    return engine.process(
        filename=filename,
        file_bytes=file_bytes,
        source_lang=source_lang,
        target_lang=target_lang,
        translate_one=translate_one,
        progress_cb=_progress_handler(progress_bar),
    )


def main() -> None:
    st.set_page_config(page_title="MMS Mission TranslatorEngine", layout="wide")
    st.title("MMS Mission TranslatorEngine")
    st.caption("DOCX/PPTX 翻译（保留原始格式）")

    laozhang_api_key, base_url = load_laozhang_config()
    if not laozhang_api_key:
        laozhang_api_key = st.secrets.get("LAOZHANG_API_KEY")
    if not base_url:
        base_url = "https://api.laozhang.ai/v1"

    gemini_api_key = load_gemini_key() or st.secrets.get("GEMINI_API_KEY")

    uploaded_file = st.file_uploader(
        "上传文件",
        type=["docx", "pptx"],
    )

    provider = st.selectbox("模型供应商", PROVIDER_OPTIONS, index=0)

    if provider == "Gemini(官方)":
        gemini_api_key = st.text_input(
            "Gemini API Key（可覆盖 .env / Secrets）",
            value=gemini_api_key or "",
            type="password",
        )

    col1, col2, col3 = st.columns(3)
    with col1:
        source_lang = st.selectbox("源语言", LANG_OPTIONS, index=0)
    with col2:
        target_lang = st.selectbox("目标语言", LANG_OPTIONS, index=1)
    with col3:
        model_candidates = GEMINI_MODELS if provider == "Gemini(官方)" else LAOZHANG_MODELS
        model_name = st.selectbox("模型", model_candidates, index=0)

    topic_hint = st.text_input("文档主题（可选）", value="")

    translate_clicked = st.button("开始翻译", type="primary", disabled=not uploaded_file)
    if translate_clicked:
        if provider == "Gemini(官方)" and not gemini_api_key:
            st.error("缺少 Gemini API Key，无法调用模型。")
            return
        if provider != "Gemini(官方)" and not laozhang_api_key:
            st.error("缺少 LAOZHANG_API_KEY，无法调用模型。")
            return

        progress_bar = st.progress(0.0)
        with st.spinner("翻译中，请稍候..."):
            output = _translate_file(
                filename=uploaded_file.name,
                file_bytes=uploaded_file.getvalue(),
                source_lang=source_lang,
                target_lang=target_lang,
                topic_hint=topic_hint.strip(),
                provider=provider,
                model_name=model_name,
                laozhang_api_key=laozhang_api_key,
                base_url=base_url,
                gemini_api_key=gemini_api_key,
                progress_bar=progress_bar,
            )
        progress_bar.progress(1.0)
        output_name = _build_output_name(uploaded_file.name, target_lang)
        st.success("翻译完成。")
        st.download_button(
            "下载翻译后的文件",
            data=output.getvalue(),
            file_name=output_name,
        )


def _build_output_name(filename: str, target_lang: str) -> str:
    if "." not in filename:
        return f"{filename}_{target_lang}"
    base, ext = filename.rsplit(".", 1)
    return f"{base}_{target_lang}.{ext}"


if __name__ == "__main__":
    main()

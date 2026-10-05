import json
import os
from urllib.parse import quote

from doc_builder import create_test_docx
from dotenv import load_dotenv
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from google import genai
from google.genai import types
from ppt_builder import create_presentation_pptx
from pydantic import BaseModel, Field

load_dotenv("zalupa.env")

api_key = os.getenv("GEMINI_API_KEY")
app = FastAPI(title="Luxury Teacher Assistant")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class TestRequest(BaseModel):
    subject: str = Field(..., description="Предмет")
    topic: str = Field(..., description="Тема урока")
    grade: str = Field(..., description="Класс")
    questions_count: int = Field(5, description="Количество вопросов")
    test_type: str = Field("anti_cheat", description="Тип теста")
    question_format: str = Field(
        "single", description="single | multiple | mixed"
    )
    custom_notes: str = Field("", description="Пожелания")


class PresentationRequest(BaseModel):
    subject: str = Field(..., description="Предмет")
    topic: str = Field(..., description="Тема урока")
    grade: str = Field(..., description="Класс")
    slides_count: int = Field(6, description="Количество слайдов")
    custom_notes: str = Field("", description="Пожелания")


@app.get("/")
def read_root():
    return {"message": "Luxury Teacher Assistant Backend Running!"}


@app.post("/generate-test")
def generate_test(req: TestRequest):
    if not api_key:
        raise HTTPException(status_code=500, detail="API ключ не найден.")

    client = genai.Client(api_key=api_key)

    mode_instruction = (
        "РЕЖИМ: Обычный проверочный тест."
        if req.test_type == "standard"
        else "РЕЖИМ: Анти-списываемый тест (на логику и подтекст)."
    )

    format_instruction = ""
    if req.question_format == "multiple":
        format_instruction = "ФОРМАТ: В каждом вопросе должно быть СТРОГО 2 ИЛИ БОЛЕЕ правильных ответов."
    elif req.question_format == "mixed":
        format_instruction = "ФОРМАТ: Смешанная работа. Часть вопросов — обычный тест, а часть (минимум 2) — задания на СООТВЕТСТВИЕ (type = 'matching')."
    else:
        format_instruction = (
            "ФОРМАТ: Классический тест с 1 правильным ответом."
        )

    system_instruction = f"""
    Ты — экспертный методист по литературе и русскому языку.
    {mode_instruction}
    {format_instruction}
    
    Верни СТРОГО JSON со следующей структурой без markdown:
    {{
      "title": "Название работы",
      "subject": "Предмет",
      "grade": "Класс",
      "questions": [
        {{
          "id": 1,
          "type": "single",  // "single" (1 ответ), "multiple" (несколько ответов) или "matching" (соответствие)
          "question": "Текст вопроса или инструкция",
          "options": ["А) Вариант 1", "Б) Вариант 2", "В) Вариант 3", "Г) Вариант 4"],
          "correct_answers": ["Б) Вариант 2"], // массив верных строк из options
          "explanation": "Подробное объяснение для учителя",
          // Точно заполни поля ниже ТОЛЬКО если type == "matching":
          "left_column": ["1. Элемент A", "2. Элемент B"],
          "right_column": ["А) Описание A", "Б) Описание B"],
          "correct_pairs": ["1 — А", "2 — Б"]
        }}
      ]
    }}
    """

    user_prompt = f"Предмет: {req.subject}, Тема: {req.topic}, Класс: {req.grade}, Вопросов: {req.questions_count}, Пожелания: {req.custom_notes}"

    try:
        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=user_prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                response_mime_type="application/json",
                temperature=0.7,
            ),
        )
        return {"success": True, "data": json.loads(response.text)}
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Ошибка генерации: {str(e)}"
        )


@app.post("/download-docx")
def download_docx(payload: dict):
    try:
        test_data = payload.get("data", payload)
        doc_mode = payload.get("doc_mode", "full")
        docx_stream = create_test_docx(test_data, doc_mode=doc_mode)

        prefix = (
            "Тест"
            if doc_mode == "student"
            else ("Ответы" if doc_mode == "teacher" else "Полный_тест")
        )
        filename = f"{prefix}_{test_data.get('subject', 'Задание')}.docx"
        encoded_filename = quote(filename)

        return StreamingResponse(
            docx_stream,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
            },
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка Word: {str(e)}")


@app.post("/generate-presentation")
def generate_presentation(req: PresentationRequest):
    if not api_key:
        raise HTTPException(status_code=500, detail="API ключ не найден.")

    client = genai.Client(api_key=api_key)

    system_instruction = """
    Ты — учитель-методист. Составь структуру презентации.
    Верни СТРОГО JSON без markdown:
    {
      "title": "Тема",
      "subject": "Предмет",
      "grade": "Класс",
      "slides": [
        {
          "slide_number": 1,
          "title": "Заголовок слайда",
          "content_bullets": ["Тезис 1", "Тезис 2"],
          "teacher_notes": "Заметка учителя"
        }
      ]
    }
    """

    user_prompt = f"Предмет: {req.subject}, Тема: {req.topic}, Класс: {req.grade}, Слайдов: {req.slides_count}, Пожелания: {req.custom_notes}"

    try:
        response = client.models.generate_content(
            model="gemini-3.5-flash-lite",
            contents=user_prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_instruction,
                response_mime_type="application/json",
                temperature=0.7,
            ),
        )
        return {"success": True, "data": json.loads(response.text)}
    except Exception as e:
        raise HTTPException(
            status_code=500, detail=f"Ошибка генерации: {str(e)}"
        )


@app.post("/download-pptx")
def download_pptx(payload: dict):
    try:
        ppt_data = payload.get("data", payload)
        pptx_stream = create_presentation_pptx(ppt_data)
        filename = f"Презентация_{ppt_data.get('subject', 'Урок')}.pptx"
        encoded_filename = quote(filename)

        return StreamingResponse(
            pptx_stream,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={
                "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
            },
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Ошибка PPTX: {str(e)}")
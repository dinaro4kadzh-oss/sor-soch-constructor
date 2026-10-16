from io import BytesIO
import json
import re
import os

# Важно: системные сертификаты Windows подключаются до импорта Google SDK.
try:
    import truststore
    truststore.inject_into_ssl()
except ImportError:
    pass

import streamlit as st
from docx import Document
from docx.shared import Pt
from pypdf import PdfReader
from dotenv import load_dotenv
from google import genai
from google.genai import types

load_dotenv()

MODEL_NAME = os.getenv("GEMINI_MODEL", "gemini-3.5-flash-lite")
MAX_SOURCE_CHARS = 85000

st.set_page_config(page_title="Конструктор СОР и СОЧ", page_icon="📘", layout="wide")
st.title("📘 Конструктор СОР и СОЧ")
st.write(
    "Инструмент для учителей информатики 5–11 классов Казахстана. "
    "Загрузите учебный материал: ИИ подготовит задания, ответы, критерии и дескрипторы."
)

TASK_TYPES = [
    "Один правильный ответ",
    "Несколько правильных ответов",
    "Установление соответствия",
    "Вычисление или решение",
    "Краткий ответ",
    "Развёрнутый ответ",
    "Практическое задание",
]


def read_uploaded_file(uploaded_file):
    """Извлекает текст из DOCX, PDF или TXT."""
    name = uploaded_file.name.lower()
    try:
        if name.endswith(".docx"):
            doc = Document(uploaded_file)
            parts = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells]
                    if any(cells):
                        parts.append(" | ".join(cells))
            return "\n".join(parts)
        if name.endswith(".pdf"):
            reader = PdfReader(uploaded_file)
            return "\n\n".join(
                f"[Страница {i + 1}]\n{page.extract_text() or ''}"
                for i, page in enumerate(reader.pages)
            )
        if name.endswith(".txt"):
            data = uploaded_file.getvalue()
            for encoding in ("utf-8-sig", "cp1251"):
                try:
                    return data.decode(encoding)
                except UnicodeDecodeError:
                    pass
            return data.decode("utf-8", errors="replace")
    except Exception as error:
        return f"ОШИБКА ЧТЕНИЯ ФАЙЛА: {error}"
    return ""

def get_gemini_client():
    api_key = (
        st.secrets.get("GEMINI_API_KEY", "")
        if hasattr(st, "secrets")
        else ""
    ) or os.getenv("GEMINI_API_KEY", "")
    api_key = api_key.strip()

    if not api_key:
        raise RuntimeError(
            "Не найден GEMINI_API_KEY. Проверьте настройки Secrets в "
            "Streamlit Cloud или наличие ключа в файле .env."
        )
    return genai.Client(api_key=api_key)

def extract_json(text):
    """Получает JSON даже если модель обрамит ответ markdown-блоком."""
    cleaned = (text or "").strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        start = cleaned.find("{")
        end = cleaned.rfind("}")
        if start >= 0 and end > start:
            return json.loads(cleaned[start:end + 1])
        raise ValueError("ИИ вернул ответ не в формате JSON. Попробуйте сформировать работу ещё раз.")


def clean_option_text(option, index):
    """Удаляет случайные метки A), B) и их дубли из текста варианта."""
    text = str(option).strip()
    expected_label = chr(65 + index)
    # Модель иногда возвращает "A) A) текст", хотя буква добавляется приложением.
    for _ in range(3):
        match = re.match(r"^\s*([A-F])[).:\-]\s*", text, flags=re.IGNORECASE)
        if not match or match.group(1).upper() != expected_label:
            break
        text = text[match.end():].strip()
    return text


def get_leading_answer_letters(answer):
    """Извлекает только буквы ключа в начале строки, не буквы из пояснения."""
    text = str(answer).strip()
    text = re.sub(r"^правильные\s+ответы?\s*:?\s*", "", text, flags=re.IGNORECASE)
    match = re.match(
        r"^([A-F](?:\s*(?:,|;|/|и|&)\s*[A-F])*)(?=\s*(?:[.):\-]|$|\s))",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return []
    return [letter.upper() for letter in re.findall(r"[A-F]", match.group(1), flags=re.IGNORECASE)]


def normalize_task(task, index, settings, objective):
    """Проверяет обязательные поля и приводит ответ модели к единому виду."""
    if not isinstance(task, dict):
        raise ValueError(f"Задание {index}: неверный формат ответа ИИ.")

    question = str(task.get("question", "")).strip()
    answer = str(task.get("correct_answer", "")).strip()
    if len(question) < 12:
        raise ValueError(f"Задание {index}: ИИ не сформировал полноценное условие.")
    if not answer:
        raise ValueError(f"Задание {index}: не сформирован ключ ответа.")

    options = task.get("options", [])
    if not isinstance(options, list):
        options = []
    options = [clean_option_text(x, i) for i, x in enumerate(options) if str(x).strip()]

    matching_left = task.get("matching_left", [])
    matching_right = task.get("matching_right", [])
    if not isinstance(matching_left, list):
        matching_left = []
    if not isinstance(matching_right, list):
        matching_right = []
    matching_left = [str(x).strip() for x in matching_left if str(x).strip()]
    matching_right = [str(x).strip() for x in matching_right if str(x).strip()]

    criteria = task.get("criteria", [])
    descriptors = task.get("descriptors", [])
    if isinstance(criteria, str):
        criteria = [criteria]
    if isinstance(descriptors, str):
        descriptors = [descriptors]
    criteria = [str(x).strip() for x in criteria if str(x).strip()] if isinstance(criteria, list) else []
    descriptors = [str(x).strip() for x in descriptors if str(x).strip()] if isinstance(descriptors, list) else []

    task_type = settings["type"]
    if task_type == "Один правильный ответ" and len(options) != 4:
        raise ValueError(f"Задание {index}: для одного правильного ответа нужно ровно 4 варианта.")
    if task_type == "Несколько правильных ответов" and not 3 <= len(options) <= 6:
        raise ValueError(f"Задание {index}: нужно от 3 до 6 вариантов ответа.")
    if task_type == "Несколько правильных ответов":
        correct_letters = get_leading_answer_letters(answer)
        unique_letters = set(correct_letters)
        allowed_letters = {chr(65 + i) for i in range(len(options))}
        if len(unique_letters) < 2:
            raise ValueError(
                f"Задание {index}: в ключе для задания с несколькими правильными ответами "
                "укажите минимум две буквы в начале, например «A, C. Пояснение…»."
            )
        if not unique_letters.issubset(allowed_letters):
            raise ValueError(
                f"Задание {index}: в ключе указаны буквы, которым не соответствуют варианты ответа. "
                f"Допустимые буквы: {', '.join(sorted(allowed_letters))}."
            )
    if task_type == "Один правильный ответ":
        correct_markers = re.findall(r"(?:^|[;\n,])\s*([A-D])\s*[).:-]", answer, flags=re.IGNORECASE)
        if len(set(x.upper() for x in correct_markers)) > 1:
            raise ValueError(f"Задание {index}: тип предполагает один правильный ответ, но ключ, похоже, содержит несколько вариантов.")
    if task_type == "Установление соответствия":
        if len(matching_left) < 2 or len(matching_right) < 2:
            raise ValueError(
                f"Задание {index}: для установления соответствия нужны две заполненные группы "
                "элементов (не менее двух элементов в каждой). ИИ должен вернуть matching_left и matching_right."
            )
        if len(set(matching_left)) != len(matching_left) or len(set(matching_right)) != len(matching_right):
            raise ValueError(f"Задание {index}: в группах для соответствия обнаружены повторяющиеся элементы.")
        if len(matching_left) != len(matching_right):
            raise ValueError(f"Задание {index}: в обеих группах должно быть одинаковое количество элементов. Сейчас: {len(matching_left)} и {len(matching_right)}.")
        # Не требуем буквального совпадения всех формулировок: модель может
        # корректно описать пару, сокращая длинные названия элементов.
        # Наличие обеих групп и непустого ключа проверяется выше; точность пар
        # необходимо проверить в итоговой схеме оценивания.
        if len(answer.strip()) < 12:
            raise ValueError(
                f"Задание {index}: ключ соответствия слишком короткий. "
                "Укажите пару для каждого элемента группы 1."
            )
    if not criteria:
        raise ValueError(f"Задание {index}: не сформированы конкретные критерии оценивания.")
    if not descriptors:
        raise ValueError(f"Задание {index}: не сформированы конкретные дескрипторы.")

    # Проверяем баллы, если они указаны в дескрипторах в формате «— 1 балл».
    descriptor_scores = []
    for descriptor in descriptors:
        matches = re.findall(r"(?:—|-|:)\s*(\d+)\s*балл", descriptor, flags=re.IGNORECASE)
        descriptor_scores.extend(int(value) for value in matches)
    if not descriptor_scores:
        raise ValueError(
            f"Задание {index}: в дескрипторах не найдены баллы. Укажите баллы явно, например «Определяет адрес сети — 1 балл»."
        )
    if sum(descriptor_scores) != int(settings["points"]):
        raise ValueError(
            f"Задание {index}: сумма баллов в дескрипторах равна {sum(descriptor_scores)}, "
            f"а заданию выделено {settings['points']} балл(а/ов). Исправьте дескрипторы."
        )

    if not str(task.get("source_basis", "")).strip():
        raise ValueError(f"Задание {index}: не указано, как задание связано с загруженным материалом.")

    return {
        "number": index,
        "type": task_type,
        "objective": objective,
        "points": int(settings["points"]),
        "question": question,
        "options": options,
        "matching_left": matching_left,
        "matching_right": matching_right,
        "correct_answer": answer,
        "explanation": str(task.get("explanation", "")).strip(),
        "criteria": criteria,
        "descriptors": descriptors,
        "source_basis": str(task.get("source_basis", "")).strip(),
    }


def generate_assessment(grade, assessment, quarter, max_score, objectives, source_text, settings):
    if not source_text.strip():
        raise ValueError("Не удалось извлечь текст из учебного материала. Если это скан PDF, нужен PDF с распознанным текстом или DOCX/TXT.")

    source_for_prompt = source_text[:MAX_SOURCE_CHARS]
    if len(source_text) > MAX_SOURCE_CHARS:
        st.warning(f"Материал большой: для генерации использованы первые {MAX_SOURCE_CHARS:,} символов текста. Для более точной работы загрузите нужную главу отдельно.")

    objectives_for_tasks = [item.get("objective") or objectives[i % len(objectives)] for i, item in enumerate(settings)]
    task_specs = [
        {"number": i + 1, "type": item["type"], "points": item["points"], "learning_objective": objectives_for_tasks[i]}
        for i, item in enumerate(settings)
    ]

    prompt = f"""
Ты — опытный методист по информатике в школе Казахстана. Создай ГОТОВУЮ суммативную оценочную работу, опираясь прежде всего на предоставленный учебный материал.

Параметры:
- Вид работы: {assessment}
- Класс: {grade}
- Четверть: {quarter}
- Максимальный балл: {max_score}
- Структура заданий (строго соблюдай тип и баллы):
{json.dumps(task_specs, ensure_ascii=False, indent=2)}

Цели обучения:
{json.dumps(objectives, ensure_ascii=False, indent=2)}

УЧЕБНЫЙ МАТЕРИАЛ (это основной источник фактов):
--- НАЧАЛО МАТЕРИАЛА ---
{source_for_prompt}
--- КОНЕЦ МАТЕРИАЛА ---

Обязательные требования:
1. Создай каждое задание целиком. Не пиши инструкции учителю вида «составьте вопрос», «предложите установить соответствие», «заполните ключ».
2. Каждое задание должно проверять именно указанную для него цель обучения и быть посильным для указанного класса. Перед выдачей результата отдельно проверь: какое конкретное действие ученика требуется в цели и действительно ли это действие выполняется в задании. Не подменяй вычисление адреса сети вопросом на определение термина; не объединяй разные цели в одно задание без прямого указания.
3. Используй термины, факты, алгоритмы и примеры из учебного материала. Не цитируй случайные обрывки текста вместо задания. Поле source_basis кратко указывает, какая идея материала использована.
4. Не утверждай, что материал содержит сведения, которых в нём нет. Если выбранная цель обучения не раскрыта в источнике, явно напиши это в source_basis и не выдумывай содержание от имени источника. Допустимо использовать общеизвестные знания только для корректного задания, но обозначь, что материал источника недостаточен.
5. Для «Один правильный ответ» создай ровно 4 осмысленных варианта и один верный. В поле options пиши только текст вариантов, без буквенных меток A), B) и т. п. Метки приложение добавит самостоятельно.
6. Для «Несколько правильных ответов» создай 4–6 вариантов; правильных ответов должно быть минимум два. В поле options пиши только текст вариантов, без буквенных меток A), B) и т. п. Метки приложение добавит самостоятельно. В поле correct_answer ОБЯЗАТЕЛЬНО сначала укажи только буквы всех правильных вариантов в формате «A, C» или «A, C, E», используя реальные метки вариантов A, B, C, D, E, F по порядку options. После букв можно добавить точку и пояснение, например «A, C. Эти варианты верны, потому что ...». Не заменяй буквы описанием ответа и не пиши номера вместо букв.
7. Для «Установление соответствия» ОБЯЗАТЕЛЬНО создай две отдельные непустые группы: matching_left и matching_right, минимум по 2 элемента в каждой. Не прячь элементы только в correct_answer или explanation. В question напиши краткую инструкцию. Все элементы должны быть показаны ученику в документе. В correct_answer дай полное соответствие для каждого элемента из matching_left с соответствующим элементом из matching_right. Для ясности используй формат «[элемент группы 1] — [элемент группы 2]» отдельной строкой для каждой пары. Можно сокращать длинные формулировки только так, чтобы однозначно узнавался исходный элемент. Количество элементов в группах должно быть согласованным; не добавляй отсутствующие элементы в ключ.
8. Для «Вычисление или решение» создай задачу с конкретными исходными данными и корректным пошаговым решением в correct_answer.
9. Для «Краткий ответ» вопрос должен иметь однозначный ожидаемый ответ.
10. Для «Развёрнутый ответ» потребуй объяснение/обоснование; в ключе перечисли существенные элементы ответа.
11. Для «Практическое задание» дай конкретные действия и проверяемый ожидаемый результат. Не требуй программ или программного обеспечения, не соответствующих материалу.
12. Критерии и дескрипторы должны быть конкретными и оценивать действия из задания. Для каждого задания сумма баллов, указанных в дескрипторах, должна равняться баллам, выделенным именно на это задание, а не максимальному баллу всей работы. Каждый дескриптор с баллами пиши в явном формате «Действие учащегося — 1 балл».
13. Распределение баллов по заданиям фиксировано и в сумме равно {max_score}. Не меняй его.
14. Проверь правильность ключей, вычислений, вариантов и соответствие цели обучения перед ответом. Для вычислений пересчитай результат независимо и покажи единицы на каждом шаге. Для заданий с выбором проверь, что ключ точно указывает правильные варианты и не включает неверные. Для соответствия ключ должен содержать каждую пару и не вводить элементы, отсутствующие в группах. Если не можешь подтвердить факт по источнику, не приписывай его источнику: прямо укажи ограничение в source_basis.
15. Ответ верни только в корректном JSON без markdown и пояснений вне JSON, по схеме:
{{
  "tasks": [
    {{
      "question": "полное условие задания",
      "options": ["вариант A", "вариант B", "вариант C", "вариант D"],
      "matching_left": ["элемент 1", "элемент 2"],
      "matching_right": ["описание А", "описание Б"],
      "correct_answer": "правильный ответ или полное решение",
      "explanation": "пояснение ключа для учителя",
      "criteria": ["критерий оценивания"],
      "descriptors": ["действие учащегося — N балл(а/ов)"],
      "source_basis": "кратко: какая часть/идея учебного материала использована"
    }}
  ]
}}
Для типов, где варианты не нужны, верни пустой массив options. Для всех типов, кроме «Установление соответствия», верни пустые массивы matching_left и matching_right. Для типа «Установление соответствия» заполни оба массива реальными элементами, которые должны быть видны ученику. Верни ровно {len(settings)} объектов tasks в исходном порядке.
""".strip()

    client = get_gemini_client()
    last_error = None
    for attempt in range(3):
        try:
            if attempt == 0:
                request_prompt = prompt
            else:
                matching_numbers = [
                    i + 1 for i, item in enumerate(settings)
                    if item["type"] == "Установление соответствия"
                ]
                correction = (
                    "\n\nВАЖНО: предыдущий ответ не прошёл автоматическую проверку. "
                    + f"Причина: {last_error}. Исправь именно эту ошибку. "
                    + "Для каждого задания типа «Несколько правильных ответов» поле correct_answer должно начинаться с букв правильных вариантов в формате «A, C» или «A, C, E»; буквы должны соответствовать options. В options возвращай только текст вариантов без меток A), B) и т. п. "
                )
                if matching_numbers:
                    correction += (
                        "ВНИМАНИЕ ДЛЯ ЗАДАНИЙ НА СООТВЕТСТВИЕ: номера заданий "
                        + ", ".join(map(str, matching_numbers))
                        + " обязательно должны иметь заполненные массивы matching_left и matching_right, минимум по 2 непустых текстовых элемента в каждом. "
                        + "Не возвращай для них пустые массивы и не помещай элементы только в question или correct_answer. "
                        + "Пример структуры: \"matching_left\":[\"IP-адрес узла\",\"Маска подсети\"], \"matching_right\":[\"Адрес устройства в сети\",\"Определяет сетевую и узловую части адреса\"]. "
                        + "В correct_answer перечисли каждую пару в формате «элемент группы 1 — элемент группы 2». "
                    )
                request_prompt = prompt + correction + " Верни только исправленный JSON по исходной схеме, без комментариев."
            response = client.models.generate_content(
                model=MODEL_NAME,
                contents=request_prompt,
                config=types.GenerateContentConfig(
                    temperature=0.25,
                    response_mime_type="application/json",
                ),
            )
            payload = extract_json(response.text or "")
            raw_tasks = payload.get("tasks") if isinstance(payload, dict) else None
            if not isinstance(raw_tasks, list) or len(raw_tasks) != len(settings):
                raise ValueError(f"Ожидалось {len(settings)} заданий, получено: {len(raw_tasks) if isinstance(raw_tasks, list) else 'не список'}. Генерация отклонена, чтобы не выдать неполную работу.")
            tasks = [
                normalize_task(raw_tasks[i], i + 1, settings[i], objectives_for_tasks[i])
                for i in range(len(settings))
            ]
            if sum(t["points"] for t in tasks) != max_score:
                raise ValueError("Сумма баллов в результате не совпадает с заданным максимумом.")
            return tasks
        except Exception as error:
            last_error = error
    raise RuntimeError(
        f"Не удалось получить корректные задания от Gemini. Подробности: {last_error}. "
        "Проверьте интернет, доступность модели и корректность API-ключа в .env."
    )


def add_task_to_doc(doc, task, include_teacher_parts=False):
    doc.add_heading(f"Задание {task['number']} — {task['points']} балл(а/ов)", level=2)
    doc.add_paragraph(f"Цель обучения: {task['objective']}")
    doc.add_paragraph(task["question"])
    if task["options"]:
        for i, option in enumerate(task["options"]):
            label = chr(65 + i)
            doc.add_paragraph(f"{label}) {option}", style="List Bullet")
    if task["type"] == "Установление соответствия":
        doc.add_paragraph("Группа 1:")
        for i, item in enumerate(task["matching_left"], 1):
            doc.add_paragraph(f"{i}. {item}", style="List Bullet")
        doc.add_paragraph("Группа 2:")
        for i, item in enumerate(task["matching_right"]):
            doc.add_paragraph(f"{chr(65 + i)}. {item}", style="List Bullet")
    if include_teacher_parts:
        doc.add_paragraph("Критерии оценивания:")
        for item in task["criteria"]:
            doc.add_paragraph(item, style="List Bullet")
        doc.add_paragraph("Дескрипторы:")
        for item in task["descriptors"]:
            doc.add_paragraph(item, style="List Bullet")


def create_word_document(grade, assessment, quarter, max_score, objectives, sources, tasks):
    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)

    doc.add_heading(f"СУММАТИВНОЕ ОЦЕНИВАНИЕ {'ЗА РАЗДЕЛ' if assessment == 'СОР' else 'ЗА ЧЕТВЕРТЬ'}", level=0)
    doc.add_paragraph("Предмет: Информатика")
    doc.add_paragraph(f"Класс: {grade}")
    doc.add_paragraph(f"Четверть: {quarter}")
    doc.add_paragraph(f"Максимальный балл: {max_score}")
    doc.add_heading("Цели обучения", level=1)
    for objective in objectives:
        doc.add_paragraph(objective, style="List Bullet")
    doc.add_heading("Использованные материалы", level=1)
    for source in sources:
        doc.add_paragraph(source, style="List Bullet")

    doc.add_heading("Задания для учащегося", level=1)
    for task in tasks:
        add_task_to_doc(doc, task, include_teacher_parts=False)

    doc.add_page_break()
    doc.add_heading("Схема выставления баллов", level=0)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    table.autofit = False
    # Сохраняем таблицу в пределах ширины страницы A4 в книжной ориентации.
    from docx.shared import Inches
    column_widths = [Inches(0.55), Inches(2.2), Inches(0.5), Inches(2.8)]
    headers = ["№ задания", "Ответ / решение", "Баллы", "Дополнительная информация"]
    for cell, heading, width in zip(table.rows[0].cells, headers, column_widths):
        cell.text = heading
        cell.width = width
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
                run.font.name = "Arial"
                run.font.size = Pt(8)
    for task in tasks:
        cells = table.add_row().cells
        for cell, width in zip(cells, column_widths):
            cell.width = width
        cells[0].text = str(task["number"])
        answer = task["correct_answer"]
        if task["explanation"]:
            answer += "\n\nПояснение: " + task["explanation"]
        cells[1].text = answer
        cells[2].text = str(task["points"])
        details = "Критерии:\n" + "\n".join(task["criteria"])
        details += "\n\nДескрипторы:\n" + "\n".join(task["descriptors"])
        if task["source_basis"]:
            details += "\n\nОснова в источнике: " + task["source_basis"]
        cells[3].text = details
        for cell in cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(2)
                for run in paragraph.runs:
                    run.font.name = "Arial"
                    run.font.size = Pt(8)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


st.subheader("1. Параметры оценочной работы")
col1, col2 = st.columns(2)
with col1:
    grade = st.selectbox("Класс", list(range(5, 12)), format_func=lambda x: f"{x} класс")
    assessment = st.radio("Вид оценивания", ["СОР", "СОЧ"], horizontal=True)
with col2:
    quarter = st.selectbox("Четверть", ["I", "II", "III", "IV"])
    if assessment == "СОР":
        max_score = st.selectbox("Максимальный балл СОР", list(range(7, 21)), index=3)
    else:
        max_score = 25
        st.info("Максимальный балл СОЧ: 25")

st.subheader("2. Цели обучения")
objectives_text = st.text_area(
    "Введите цели обучения — каждую с новой строки",
    placeholder="Например:\n9.1.2.1 Описывать назначение сетевых протоколов\n9.1.2.2 Определять адрес сети по IP-адресу и маске",
    height=130,
)
objectives = [line.strip() for line in objectives_text.splitlines() if line.strip()]
objective_choices = ["— выберите цель обучения —"] + objectives

# Предупреждаем о возможном несоответствии класса и кода цели обучения.
# Для типовых кодов Казахстана первая цифра после первой точки обычно указывает ступень/класс.
objective_grade_mismatches = []
for objective in objectives:
    match = re.match(r"^\s*(\d{1,2})\.(\d+)\.", objective)
    if match:
        code_grade = int(match.group(1))
        if code_grade != int(grade):
            objective_grade_mismatches.append((objective, code_grade))
if objective_grade_mismatches:
    st.warning(
        "Возможное несоответствие класса и целей обучения. Выбран "
        f"{grade} класс, но обнаружены коды целей другого класса: "
        + "; ".join(f"{obj} (код начинается с {code_grade})" for obj, code_grade in objective_grade_mismatches)
        + ". Проверьте ввод перед генерацией. Это предупреждение, так как формат кода может различаться."
    )

st.subheader("3. Учебные материалы")
uploaded_files = st.file_uploader(
    "Загрузите учебник, конспект или другие материалы",
    type=["docx", "pdf", "txt"],
    accept_multiple_files=True,
)

st.subheader("4. Структура работы")
number_of_tasks = st.number_input("Количество заданий", min_value=1, max_value=15, value=5, step=1, key="number_of_tasks")
st.caption("Количество блоков ниже автоматически меняется при выборе числа заданий.")
task_settings = []
for i in range(int(number_of_tasks)):
    with st.container(border=True):
        st.markdown(f"**Задание {i + 1}**")
        task_col1, task_col2 = st.columns(2)
        with task_col1:
            task_type = st.selectbox("Формат задания", TASK_TYPES, key=f"type_{i}")
            if objectives:
                objective_index = st.selectbox(
                    "Цель обучения для этого задания",
                    options=list(range(len(objective_choices))),
                    format_func=lambda j, choices=objective_choices: choices[j],
                    index=min(i + 1, len(objective_choices) - 1),
                    key=f"objective_{i}",
                )
                selected_objective = objectives[objective_index - 1] if objective_index > 0 else ""
            else:
                selected_objective = ""
                st.info("Сначала введите цели обучения выше.")
        with task_col2:
            points = st.number_input("Баллы за задание", min_value=1, max_value=25, value=1, step=1, key=f"points_{i}")
        task_settings.append({"type": task_type, "points": int(points), "objective": selected_objective})

total_points = sum(task["points"] for task in task_settings)
if total_points == max_score:
    st.success(f"Сумма баллов: {total_points} из {max_score} — всё сходится.")
else:
    st.warning(f"Сумма баллов сейчас {total_points}, а максимум работы — {max_score}. Измените баллы заданий перед генерацией.")

submitted = st.button("✨ Сгенерировать задания с помощью ИИ", type="primary")

if submitted:
    objectives = [line.strip() for line in objectives_text.splitlines() if line.strip()]
    if not objectives:
        st.error("Введите хотя бы одну цель обучения.")
        st.stop()
    if not uploaded_files:
        st.error("Загрузите учебный материал, на основе которого нужно составить задания.")
        st.stop()
    total_points = sum(task["points"] for task in task_settings)
    if total_points != max_score:
        st.error(f"Сумма баллов за задания равна {total_points}, а максимум работы — {max_score}. Измените баллы заданий.")
        st.stop()

    source_texts = []
    source_names = []
    for uploaded_file in uploaded_files:
        text = read_uploaded_file(uploaded_file)
        if text.startswith("ОШИБКА ЧТЕНИЯ ФАЙЛА:"):
            st.error(f"{uploaded_file.name}: {text}")
            st.stop()
        if text.strip():
            source_texts.append(text)
            source_names.append(uploaded_file.name)
    combined_text = "\n\n".join(source_texts)
    if not combined_text.strip():
        st.error("Из файла не удалось извлечь текст. Возможно, PDF состоит из сканов. Попробуйте файл с распознанным текстом.")
        st.stop()

    try:
        with st.spinner("Gemini изучает учебный материал и формирует задания, ключи и дескрипторы. Это может занять минуту..."):
            tasks = generate_assessment(grade, assessment, quarter, max_score, objectives, combined_text, task_settings)
        st.session_state["generated_tasks"] = tasks
        st.session_state["generated_info"] = {
            "grade": grade, "assessment": assessment, "quarter": quarter,
            "max_score": max_score, "objectives": objectives, "sources": source_names,
        }
        st.success("Задания сгенерированы. Проверьте содержание перед использованием в классе.")
    except Exception as error:
        st.error(f"Ошибка генерации: {error}")
       st.info("Если ошибка связана с подключением, проверьте интернет и настройку GEMINI_API_KEY в секретах Streamlit Cloud или в файле .env. Не публикуйте сам ключ.")

if "generated_tasks" in st.session_state and "generated_info" in st.session_state:
    tasks = st.session_state["generated_tasks"]
    info = st.session_state["generated_info"]
    st.divider()
    st.subheader("5. Сгенерированная оценочная работа")
    for task in tasks:
        with st.expander(f"Задание {task['number']} — {task['points']} балл(а/ов)", expanded=True):
            st.markdown(f"**Тип:** {task['type']}")
            st.markdown(f"**Цель обучения:** {task['objective']}")
            st.write(task["question"])
            if task["options"]:
                for i, option in enumerate(task["options"]):
                    st.write(f"{chr(65 + i)}) {clean_option_text(option, i)}")
            if task["type"] == "Установление соответствия":
                left_col, right_col = st.columns(2)
                with left_col:
                    st.markdown("**Группа 1**")
                    for i, item in enumerate(task["matching_left"], 1):
                        st.write(f"{i}. {item}")
                with right_col:
                    st.markdown("**Группа 2**")
                    for i, item in enumerate(task["matching_right"]):
                        st.write(f"{chr(65 + i)}. {item}")
            st.markdown("**Критерии оценивания**")
            for item in task["criteria"]:
                st.write(f"- {item}")
            st.markdown("**Дескрипторы**")
            for item in task["descriptors"]:
                st.write(f"- {item}")
            with st.expander("Ответ и пояснение для учителя"):
                st.write(task["correct_answer"])
                if task["explanation"]:
                    st.write(task["explanation"])
                if task["source_basis"]:
                    st.caption("Связь с материалом: " + task["source_basis"])

    word_file = create_word_document(info["grade"], info["assessment"], info["quarter"], info["max_score"], info["objectives"], info["sources"], tasks)
    st.download_button(
        label="📄 Скачать СОР/СОЧ в Word",
        data=word_file,
        file_name=f"{info['assessment']}_informatika_{info['grade']}_klass.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

"""Local, deterministic AI-controller heuristics for the hackathon prototype.

These checks are intentionally transparent rules, not a hosted LLM or a safety
certification. The master retains the final decision on every work order.
"""
from __future__ import annotations

import base64
import hashlib
from datetime import datetime
from io import BytesIO
from typing import Any

try:
    from PIL import Image
except ImportError:  # Optional image comparison; exact-byte check still works.
    Image = None


def photo_similarity(before: str | None, after: str | None) -> float | None:
    if not before or not after:
        return None
    try:
        a = base64.b64decode(before.split(",")[-1])
        b = base64.b64decode(after.split(",")[-1])
        if Image is None:
            return 1.0 if hashlib.sha256(a).digest() == hashlib.sha256(b).digest() else None

        def average_hash(raw: bytes) -> list[bool]:
            pixels = list(Image.open(BytesIO(raw)).convert("L").resize((8, 8)).getdata())
            mean = sum(pixels) / len(pixels)
            return [pixel >= mean for pixel in pixels]

        hash_a, hash_b = average_hash(a), average_hash(b)
        return sum(left == right for left, right in zip(hash_a, hash_b)) / 64
    except Exception:
        return None


def evaluate_order(order: dict[str, Any]) -> dict[str, Any]:
    """Return deterministic completeness, time, material, and photo signals."""
    issues: list[str] = []
    work_done = (order.get("work_done") or "").strip()
    if len(work_done) < 8:
        issues.append("Добавьте подробное описание выполненных работ.")
    if not order.get("defect_code_id"):
        issues.append("Выберите шифр неисправности.")

    materials = order.get("materials") or []
    if not materials:
        issues.append("Укажите списанные материалы или подтвердите, что материалы не использовались.")
    elif sum(float(item.get("quantity", 0) or 0) for item in materials) > 10:
        issues.append("Расход ТМЦ выше типового уровня; проверьте количество и необходимость списания.")

    if order.get("job_type") != "Плановый" and not order.get("after_photo"):
        issues.append("Для внеплановой работы требуется фото после ремонта.")

    description_tokens = {word.lower().strip(".,:;!?()") for word in (order.get("description") or "").split() if len(word) > 5}
    work_tokens = {word.lower().strip(".,:;!?()") for word in work_done.split() if len(word) > 5}
    if description_tokens and work_tokens and not description_tokens.intersection(work_tokens):
        issues.append("Описание закрытия слабо связано с первоначальной неисправностью; проверьте соответствие работ.")

    if order.get("started_at") and order.get("completed_at"):
        try:
            elapsed_hours = (datetime.fromisoformat(order["completed_at"]) - datetime.fromisoformat(order["started_at"])).total_seconds() / 3600
            if elapsed_hours > 4:
                issues.append(f"Фактическая работа заняла {elapsed_hours:.1f} ч; сверьте с нормативом 2 ч и причиной простоя.")
        except (ValueError, TypeError):
            pass

    similarity = photo_similarity(order.get("before_photo"), order.get("after_photo"))
    if similarity is not None and similarity > .97:
        issues.append("Фото до и после почти не отличаются; мастер должен проверить результат визуально.")

    score = max(2.0, min(5.0, 4.9 - .55 * len(issues)))
    verdict = "Требует доработки" if issues else "Принято с замечаниями" if score < 4.5 else "Принято"
    photo_note = (
        "Почти одинаковые изображения требуют ручной проверки мастером."
        if similarity is not None and similarity > .97
        else "Сравнение показывает визуальное изменение."
        if similarity is not None
        else "Для сравнения фото «до» не приложено или анализ недоступен."
    )
    return {
        "verdict": verdict,
        "score": round(score, 1),
        "issues": issues,
        "photo_similarity": similarity,
        "photo_note": photo_note,
        "summary": "Отчёт оценён по полноте, соответствию описанию, материалам, сроку и наличию фото. Решение ИИ — рекомендация мастеру.",
    }

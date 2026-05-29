import json
import re
import shutil
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from urllib.parse import unquote, urlparse

import pandas as pd
import requests
from flask import Flask, Response, redirect, render_template_string, request, send_file, url_for
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
IMAGES_DIR = BASE_DIR / "images"
REPORT_PATH = OUTPUT_DIR / "photos_report.xlsx"
STATE_PATH = OUTPUT_DIR / "candidates.json"
ALLOWED_EXTENSIONS = {"csv", "xlsx", "xls"}
COMMONS_API_URL = "https://commons.wikimedia.org/w/api.php"
USER_AGENT = "photo-candidate-picker/1.0 (local Flask app; Wikimedia Commons API)"
REQUIRED_COLUMNS = [
    "title",
    "search_query",
    "category",
    "exact_name",
    "must_include",
    "must_exclude",
    "notes",
]

for directory in (INPUT_DIR, OUTPUT_DIR, IMAGES_DIR):
    directory.mkdir(exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


@dataclass
class ImageCandidate:
    article_title: str
    search_query: str
    category: str
    exact_name: str
    must_include: str
    must_exclude: str
    notes: str
    image_title: str
    image_page_url: str
    source_url: str
    local_path: str
    thumb_url: str
    author: str
    license_short: str
    license_url: str
    warnings: list[str] = field(default_factory=list)


def slugify(value: str, fallback: str = "item") -> str:
    value = str(value or "").strip().lower()
    value = re.sub(r"[^a-z0-9а-яё]+", "-", value, flags=re.IGNORECASE)
    value = re.sub(r"-+", "-", value).strip("-")
    return value[:80] or fallback


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def split_terms(value: str) -> list[str]:
    if not value or pd.isna(value):
        return []
    return [term.strip().lower() for term in re.split(r"[,;|]", str(value)) if term.strip()]


def read_upload(path: Path) -> pd.DataFrame:
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        df = pd.read_excel(path)

    df.columns = [str(column).strip() for column in df.columns]
    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"Нет обязательных колонок: {', '.join(missing)}")

    return df[REQUIRED_COLUMNS].fillna("")


def commons_search(query: str, limit: int = 6) -> list[dict[str, Any]]:
    params = {
        "action": "query",
        "format": "json",
        "generator": "search",
        "gsrsearch": f"filetype:bitmap {query}",
        "gsrnamespace": 6,
        "gsrlimit": limit,
        "prop": "imageinfo|info",
        "inprop": "url",
        "iiprop": "url|extmetadata|mime",
        "iiurlwidth": 900,
        "origin": "*",
    }
    response = requests.get(
        COMMONS_API_URL,
        params=params,
        headers={"User-Agent": USER_AGENT},
        timeout=25,
    )
    response.raise_for_status()
    pages = response.json().get("query", {}).get("pages", {})
    return sorted(pages.values(), key=lambda item: item.get("index", 9999))


def clean_metadata(value: Any) -> str:
    if isinstance(value, dict):
        value = value.get("value", "")
    text = re.sub(r"<[^>]+>", " ", str(value or ""))
    return re.sub(r"\s+", " ", text).strip()


def file_extension_from_url(url: str, mime: str = "") -> str:
    suffix = Path(unquote(urlparse(url).path)).suffix.lower()
    if suffix in {".jpg", ".jpeg", ".png", ".gif", ".webp", ".tif", ".tiff"}:
        return suffix
    if "png" in mime:
        return ".png"
    if "gif" in mime:
        return ".gif"
    if "webp" in mime:
        return ".webp"
    return ".jpg"


def make_accuracy_warnings(row: dict[str, str], image_title: str, metadata_text: str) -> list[str]:
    haystack = f"{image_title} {metadata_text}".lower()
    warnings: list[str] = []

    exact_name = str(row.get("exact_name", "")).strip().lower()
    if exact_name and exact_name not in haystack:
        warnings.append(f"exact_name не найден в названии/метаданных: {row['exact_name']}")

    missing_terms = [term for term in split_terms(row.get("must_include", "")) if term not in haystack]
    if missing_terms:
        warnings.append("Не найдены must_include: " + ", ".join(missing_terms))

    excluded_terms = [term for term in split_terms(row.get("must_exclude", "")) if term in haystack]
    if excluded_terms:
        warnings.append("Найдены must_exclude: " + ", ".join(excluded_terms))

    if not warnings:
        warnings.append("Проверьте соответствие вручную перед публикацией")

    return warnings


def download_image(url: str, destination: Path) -> None:
    with requests.get(url, headers={"User-Agent": USER_AGENT}, stream=True, timeout=40) as response:
        response.raise_for_status()
        with destination.open("wb") as image_file:
            for chunk in response.iter_content(chunk_size=1024 * 64):
                if chunk:
                    image_file.write(chunk)


def build_candidates(df: pd.DataFrame) -> list[ImageCandidate]:
    candidates: list[ImageCandidate] = []

    for row_index, row in df.iterrows():
        row_data = {column: str(row[column]).strip() for column in REQUIRED_COLUMNS}
        query = row_data["search_query"] or row_data["title"]
        article_slug = slugify(row_data["title"] or query, fallback=f"row-{row_index + 1}")
        article_dir = IMAGES_DIR / article_slug
        article_dir.mkdir(parents=True, exist_ok=True)

        try:
            pages = commons_search(query)
        except requests.RequestException as exc:
            candidates.append(
                ImageCandidate(
                    article_title=row_data["title"],
                    search_query=query,
                    category=row_data["category"],
                    exact_name=row_data["exact_name"],
                    must_include=row_data["must_include"],
                    must_exclude=row_data["must_exclude"],
                    notes=row_data["notes"],
                    image_title="",
                    image_page_url="",
                    source_url="",
                    local_path="",
                    thumb_url="",
                    author="",
                    license_short="",
                    license_url="",
                    warnings=[f"Ошибка Wikimedia Commons API: {exc}"],
                )
            )
            continue

        if not pages:
            candidates.append(
                ImageCandidate(
                    article_title=row_data["title"],
                    search_query=query,
                    category=row_data["category"],
                    exact_name=row_data["exact_name"],
                    must_include=row_data["must_include"],
                    must_exclude=row_data["must_exclude"],
                    notes=row_data["notes"],
                    image_title="",
                    image_page_url="",
                    source_url="",
                    local_path="",
                    thumb_url="",
                    author="",
                    license_short="",
                    license_url="",
                    warnings=["Wikimedia Commons не вернул результатов для запроса."],
                )
            )
            continue

        for candidate_index, page in enumerate(pages, start=1):
            image_info = (page.get("imageinfo") or [{}])[0]
            source_url = image_info.get("url", "")
            thumb_url = image_info.get("thumburl", source_url)
            metadata = image_info.get("extmetadata", {})
            image_title = page.get("title", "")
            metadata_text = " ".join(clean_metadata(value) for value in metadata.values())
            extension = file_extension_from_url(source_url, image_info.get("mime", ""))
            local_file = article_dir / f"{candidate_index:02d}-{slugify(image_title, 'image')}{extension}"
            local_path = ""
            warnings = make_accuracy_warnings(row_data, image_title, metadata_text)

            if source_url:
                try:
                    download_image(thumb_url or source_url, local_file)
                    local_path = str(local_file.relative_to(BASE_DIR))
                except requests.RequestException as exc:
                    warnings.append(f"Не удалось скачать изображение: {exc}")

            candidates.append(
                ImageCandidate(
                    article_title=row_data["title"],
                    search_query=query,
                    category=row_data["category"],
                    exact_name=row_data["exact_name"],
                    must_include=row_data["must_include"],
                    must_exclude=row_data["must_exclude"],
                    notes=row_data["notes"],
                    image_title=image_title,
                    image_page_url=page.get("fullurl", ""),
                    source_url=source_url,
                    local_path=local_path,
                    thumb_url=thumb_url,
                    author=clean_metadata(metadata.get("Artist", "")) or "Не указан",
                    license_short=clean_metadata(metadata.get("LicenseShortName", "")) or "Не указана",
                    license_url=clean_metadata(metadata.get("LicenseUrl", "")),
                    warnings=warnings,
                )
            )
            time.sleep(0.1)

    save_state(candidates)
    export_report(candidates)
    return candidates


def save_state(candidates: list[ImageCandidate]) -> None:
    STATE_PATH.write_text(
        json.dumps([asdict(candidate) for candidate in candidates], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def load_state() -> list[ImageCandidate]:
    if not STATE_PATH.exists():
        return []
    data = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    return [ImageCandidate(**item) for item in data]


def export_report(candidates: list[ImageCandidate]) -> None:
    rows = []
    for candidate in candidates:
        item = asdict(candidate)
        item["warnings"] = "; ".join(candidate.warnings)
        rows.append(item)
    pd.DataFrame(rows).to_excel(REPORT_PATH, index=False)


def categories_for(candidates: list[ImageCandidate]) -> list[str]:
    return sorted({candidate.category for candidate in candidates if candidate.category})


PAGE_TEMPLATE = """
<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Photo Candidate Picker</title>
  <style>
    :root { color-scheme: light; font-family: Inter, system-ui, -apple-system, Segoe UI, sans-serif; }
    body { margin: 0; background: #f5f7fb; color: #1f2937; }
    header { background: linear-gradient(135deg, #1d4ed8, #7c3aed); color: white; padding: 34px max(28px, 6vw); }
    main { padding: 26px max(20px, 5vw) 60px; }
    .hero { max-width: 1120px; margin: 0 auto; }
    .hero h1 { margin: 0 0 8px; font-size: clamp(30px, 5vw, 56px); }
    .hero p { margin: 0; opacity: .88; font-size: 18px; }
    .panel { max-width: 1120px; margin: 22px auto; background: white; border-radius: 22px; padding: 22px; box-shadow: 0 20px 50px rgba(15,23,42,.08); }
    .upload { display: grid; gap: 12px; grid-template-columns: 1fr auto; align-items: end; }
    input, select { border: 1px solid #d1d5db; border-radius: 12px; padding: 12px; background: white; }
    label { display: grid; gap: 7px; color: #4b5563; font-weight: 700; }
    button, .button { border: 0; border-radius: 12px; padding: 12px 18px; background: #2563eb; color: white; font-weight: 800; text-decoration: none; cursor: pointer; display: inline-block; }
    .button.secondary { background: #111827; }
    .message { border-radius: 14px; padding: 14px 16px; margin-top: 16px; background: #fff7ed; color: #9a3412; }
    .toolbar { display: flex; flex-wrap: wrap; gap: 12px; align-items: center; justify-content: space-between; }
    .grid { max-width: 1400px; margin: 0 auto; display: grid; gap: 22px; grid-template-columns: repeat(auto-fill, minmax(330px, 1fr)); }
    .card { background: white; border-radius: 24px; overflow: hidden; box-shadow: 0 18px 40px rgba(15,23,42,.10); display: flex; flex-direction: column; min-height: 100%; }
    .image-wrap { background: #e5e7eb; min-height: 290px; display: grid; place-items: center; }
    .image-wrap img { width: 100%; height: 330px; object-fit: cover; display: block; }
    .card-body { padding: 18px; display: grid; gap: 12px; }
    .card h2 { font-size: 20px; line-height: 1.2; margin: 0; }
    .meta { display: grid; gap: 5px; color: #4b5563; font-size: 14px; }
    .chips { display: flex; flex-wrap: wrap; gap: 7px; }
    .chip { background: #eef2ff; color: #3730a3; border-radius: 999px; padding: 5px 9px; font-size: 12px; font-weight: 800; }
    .warning { background: #fef2f2; color: #991b1b; border-left: 4px solid #ef4444; border-radius: 10px; padding: 10px 12px; font-size: 14px; }
    .links { display: flex; flex-wrap: wrap; gap: 10px; }
    .links a { color: #1d4ed8; font-weight: 800; }
    .empty { text-align: center; color: #6b7280; padding: 36px; }
    @media (max-width: 720px) { .upload { grid-template-columns: 1fr; } .image-wrap img { height: 260px; } }
  </style>
</head>
<body>
  <header>
    <div class="hero">
      <h1>Photo Candidate Picker</h1>
      <p>Локальный подбор фото-кандидатов для статей через Wikimedia Commons.</p>
    </div>
  </header>
  <main>
    <section class="panel">
      <form class="upload" action="{{ url_for('upload') }}" method="post" enctype="multipart/form-data">
        <label>Excel/CSV с колонками title, search_query, category, exact_name, must_include, must_exclude, notes
          <input type="file" name="dataset" accept=".csv,.xlsx,.xls" required>
        </label>
        <button type="submit">Загрузить и найти фото</button>
      </form>
      {% if message %}<div class="message">{{ message }}</div>{% endif %}
    </section>

    <section class="panel toolbar">
      <form action="{{ url_for('index') }}" method="get">
        <label>Фильтр по теме
          <select name="category" onchange="this.form.submit()">
            <option value="">Все темы</option>
            {% for category in categories %}
              <option value="{{ category }}" {% if selected_category == category %}selected{% endif %}>{{ category }}</option>
            {% endfor %}
          </select>
        </label>
      </form>
      <div>
        <a class="button secondary" href="{{ url_for('export') }}">Скачать photos_report.xlsx</a>
      </div>
    </section>

    {% if candidates %}
      <section class="grid">
        {% for candidate in candidates %}
          <article class="card">
            <div class="image-wrap">
              {% if candidate.local_path %}
                <img src="{{ url_for('local_file', path=candidate.local_path) }}" alt="{{ candidate.image_title }}">
              {% else %}
                <strong>Нет локального файла</strong>
              {% endif %}
            </div>
            <div class="card-body">
              <div class="chips">
                {% if candidate.category %}<span class="chip">{{ candidate.category }}</span>{% endif %}
                {% if candidate.license_short %}<span class="chip">{{ candidate.license_short }}</span>{% endif %}
              </div>
              <h2>{{ candidate.article_title }}</h2>
              <div class="meta">
                <span><strong>Запрос:</strong> {{ candidate.search_query }}</span>
                <span><strong>Файл:</strong> {{ candidate.image_title }}</span>
                <span><strong>Автор:</strong> {{ candidate.author }}</span>
                <span><strong>Заметки:</strong> {{ candidate.notes or '—' }}</span>
              </div>
              {% for warning in candidate.warnings %}<div class="warning">{{ warning }}</div>{% endfor %}
              <div class="links">
                {% if candidate.image_page_url %}<a href="{{ candidate.image_page_url }}" target="_blank" rel="noreferrer">Страница Commons</a>{% endif %}
                {% if candidate.source_url %}<a href="{{ candidate.source_url }}" target="_blank" rel="noreferrer">Источник</a>{% endif %}
                {% if candidate.license_url %}<a href="{{ candidate.license_url }}" target="_blank" rel="noreferrer">Лицензия</a>{% endif %}
              </div>
            </div>
          </article>
        {% endfor %}
      </section>
    {% else %}
      <section class="panel empty">Загрузите Excel/CSV, чтобы увидеть галерею кандидатов.</section>
    {% endif %}
  </main>
</body>
</html>
"""


@app.route("/")
def index() -> str:
    message = request.args.get("message", "")
    selected_category = request.args.get("category", "")
    all_candidates = load_state()
    visible_candidates = [
        candidate for candidate in all_candidates if not selected_category or candidate.category == selected_category
    ]
    return render_template_string(
        PAGE_TEMPLATE,
        candidates=visible_candidates,
        categories=categories_for(all_candidates),
        selected_category=selected_category,
        message=message,
    )


@app.route("/upload", methods=["POST"])
def upload() -> Response:
    uploaded_file = request.files.get("dataset")
    if not uploaded_file or not uploaded_file.filename:
        return redirect(url_for("index", message="Выберите Excel/CSV файл."))
    if not allowed_file(uploaded_file.filename):
        return redirect(url_for("index", message="Поддерживаются только .csv, .xlsx и .xls."))

    filename = secure_filename(uploaded_file.filename)
    input_path = INPUT_DIR / filename
    uploaded_file.save(input_path)

    try:
        df = read_upload(input_path)
        candidates = build_candidates(df)
    except Exception as exc:  # noqa: BLE001 - app should show a readable browser error for local users.
        return redirect(url_for("index", message=f"Ошибка обработки файла: {exc}"))

    return redirect(url_for("index", message=f"Готово: найдено {len(candidates)} фото-кандидатов."))


@app.route("/files/<path:path>")
def local_file(path: str):
    safe_path = (BASE_DIR / path).resolve()
    if not safe_path.is_relative_to(BASE_DIR):
        return "Forbidden", 403
    return send_file(safe_path)


@app.route("/export")
def export():
    candidates = load_state()
    if not REPORT_PATH.exists() and candidates:
        export_report(candidates)
    if not REPORT_PATH.exists():
        pd.DataFrame(columns=[field.name for field in ImageCandidate.__dataclass_fields__.values()]).to_excel(
            REPORT_PATH,
            index=False,
        )
    return send_file(REPORT_PATH, as_attachment=True, download_name="photos_report.xlsx")


@app.route("/reset", methods=["POST"])
def reset() -> Response:
    if STATE_PATH.exists():
        STATE_PATH.unlink()
    if REPORT_PATH.exists():
        REPORT_PATH.unlink()
    for path in IMAGES_DIR.iterdir():
        if path.is_dir():
            shutil.rmtree(path)
    return redirect(url_for("index", message="Результаты очищены."))


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)

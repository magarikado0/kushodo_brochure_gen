from __future__ import annotations

import os
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

from . import generator


def default_template_path() -> Path:
    env_path = os.environ.get("KUSHODO_TEMPLATE_PATH")
    if env_path:
        return Path(env_path)
    return Path(__file__).resolve().parent / "templates" / "パンフ鋳型.docx"


MAX_UPLOAD_BYTES = 40 * 1024 * 1024

app = FastAPI(title="京大書道部パンフレット生成 API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)

@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/generate")
async def generate_brochure(form_file: UploadFile = File(...)) -> FileResponse:
    validate_upload(form_file, ".xlsx", "作品情報フォーム.xlsx")

    work_dir = Path(tempfile.mkdtemp(prefix="kushodo-brochure-"))
    try:
        input_path = work_dir / "作品情報フォーム.xlsx"
        output_path = work_dir / f"パンフレット_{uuid.uuid4().hex[:8]}.docx"
        list_path = work_dir / "作品一覧.txt"

        await save_upload(form_file, input_path)
        template_path = prepare_template(work_dir)

        sheets = generator.read_xlsx(input_path)
        works = generator.build_works(sheets)
        if not works:
            raise user_error("入力Excelファイルに作品データがありません。")

        generator.write_list_text(works, list_path)
        generator.write_template_docx(works, template_path, output_path)

        headers = {
            "X-Work-Count": str(len(works)),
        }

        return FileResponse(
            output_path,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            filename="パンフレット_生成.docx",
            headers=headers,
            background=BackgroundTask(shutil.rmtree, work_dir, ignore_errors=True),
        )
    except HTTPException:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise
    except generator.UserFacingError as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise user_error(sanitize_user_message(str(exc))) from exc
    except PermissionError as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        target = exc.filename or "アップロードされたファイル"
        raise user_error(f"ファイルを開けませんでした: {target}") from exc
    except Exception as exc:
        shutil.rmtree(work_dir, ignore_errors=True)
        raise HTTPException(status_code=500, detail="生成中に予期しないエラーが発生しました。") from exc


def validate_upload(upload: UploadFile, extension: str, expected_name: str) -> None:
    name = upload.filename or ""
    if not name.lower().endswith(extension):
        raise user_error(f"{expected_name} に対応する {extension} ファイルを選択してください。")


def prepare_template(work_dir: Path) -> Path:
    template_path = work_dir / "パンフ鋳型.docx"
    default_path = default_template_path()
    if not default_path.is_file():
        raise user_error("サーバーに同梱されているパンフ鋳型が見つかりません。")
    shutil.copy2(default_path, template_path)
    return template_path


async def save_upload(upload: UploadFile, path: Path) -> None:
    size = 0
    with path.open("wb") as file:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                raise user_error("アップロードできるファイルサイズは1ファイル40MBまでです。")
            file.write(chunk)
    await upload.close()


def sanitize_user_message(message: str) -> str:
    lines: list[str] = []
    for line in message.splitlines():
        if line.startswith(("ファイル:", "確認する場所:")):
            prefix, _, rest = line.partition(": ")
            name = Path(rest.strip()).name if rest.strip() else rest.strip()
            lines.append(f"{prefix}: {name}")
        else:
            lines.append(line)
    return "\n".join(lines)


def user_error(message: str) -> HTTPException:
    return HTTPException(status_code=400, detail=message)



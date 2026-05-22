from __future__ import annotations

import importlib.util
import os
import shutil
import sys
import tempfile
import uuid
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask


APP_ROOT = Path(__file__).resolve().parents[1]
REPO_ROOT = APP_ROOT.parents[1]
GENERATOR_PATH = Path(os.environ.get("KUSHODO_GENERATOR_PATH", REPO_ROOT / "tools" / "generate.py"))
MAX_UPLOAD_BYTES = 40 * 1024 * 1024

app = FastAPI(title="京大書道部パンフレット生成 API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


def load_generator() -> Any:
    spec = importlib.util.spec_from_file_location("kushodo_brochure_generator", GENERATOR_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("生成エンジンを読み込めませんでした。")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


generator = load_generator()


@app.get("/api/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.post("/api/generate")
async def generate_brochure(
    form_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
) -> FileResponse:
    validate_upload(form_file, ".xlsx", "作品情報フォーム.xlsx")
    validate_upload(template_file, ".docx", "パンフ鋳型.docx")

    work_dir = Path(tempfile.mkdtemp(prefix="kushodo-brochure-"))
    try:
        input_path = work_dir / "作品情報フォーム.xlsx"
        template_path = work_dir / "パンフ鋳型.docx"
        output_path = work_dir / f"パンフレット_{uuid.uuid4().hex[:8]}.docx"
        list_path = work_dir / "作品一覧.txt"

        await save_upload(form_file, input_path)
        await save_upload(template_file, template_path)

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
        raise user_error(str(exc)) from exc
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


async def save_upload(upload: UploadFile, path: Path) -> None:
    size = 0
    with path.open("wb") as file:
        while chunk := await upload.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_UPLOAD_BYTES:
                raise user_error("アップロードできるファイルサイズは1ファイル40MBまでです。")
            file.write(chunk)
    await upload.close()


def user_error(message: str) -> HTTPException:
    return HTTPException(status_code=400, detail=message)



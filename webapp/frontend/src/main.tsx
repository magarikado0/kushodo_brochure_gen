import React, { ChangeEvent, FormEvent, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const API_BASE = import.meta.env.VITE_API_BASE ?? "";

type UploadKind = "form" | "template";

function App() {
  const [formFile, setFormFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [message, setMessage] = useState<string>("");
  const [error, setError] = useState<string>("");

  const canSubmit = useMemo(() => Boolean(formFile && templateFile && !isGenerating), [formFile, templateFile, isGenerating]);

  const updateFile = (kind: UploadKind) => (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0] ?? null;
    if (kind === "form") {
      setFormFile(file);
    } else {
      setTemplateFile(file);
    }
    setMessage("");
    setError("");
  };

  async function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!formFile || !templateFile) {
      setError("作品情報フォームとパンフ鋳型の両方を選択してください。");
      return;
    }

    setIsGenerating(true);
    setMessage("生成しています。少しお待ちください。");
    setError("");

    try {
      const body = new FormData();
      body.append("form_file", formFile);
      body.append("template_file", templateFile);

      const response = await fetch(`${API_BASE}/api/generate`, {
        method: "POST",
        body,
      });

      if (!response.ok) {
        const detail = await readError(response);
        throw new Error(detail);
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "パンフレット_生成.docx";
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      const count = response.headers.get("X-Work-Count");
      setMessage(count ? `${count}件の作品を生成しました。` : "パンフレットを生成しました。");
    } catch (caught) {
      setError(caught instanceof Error ? caught.message : "生成に失敗しました。");
      setMessage("");
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <main className="page">
      <section className="hero">
        <div className="heroText">
          <p className="eyebrow">Kyodai Calligraphy Brochure</p>
          <h1>作品情報フォームから、パンフレットを生成します。</h1>
          <p className="lead">
            Excel と前回パンフレットの鋳型を選ぶだけで、作品一覧と作品紹介ページを流し込みます。
          </p>
        </div>
        <div className="inkMark" aria-hidden="true">書</div>
      </section>

      <form className="panel" onSubmit={submit}>
        <FilePicker
          title="作品情報フォーム"
          description="Google フォームなどから出力した xlsx ファイル"
          accept=".xlsx"
          file={formFile}
          onChange={updateFile("form")}
        />
        <FilePicker
          title="パンフ鋳型"
          description="前回パンフレットをコピーした docx ファイル"
          accept=".docx"
          file={templateFile}
          onChange={updateFile("template")}
        />

        <button className="generateButton" disabled={!canSubmit}>
          {isGenerating ? "生成中..." : "パンフレットを生成"}
        </button>

        {message && <p className="message">{message}</p>}
        {error && <p className="error">{error}</p>}
      </form>

      <section className="notes">
        <h2>生成後にすること</h2>
        <p>作品画像は自動では入りません。Word ファイル内の画像プレースホルダに手で配置してください。</p>
      </section>
    </main>
  );
}

function FilePicker({
  title,
  description,
  accept,
  file,
  onChange,
}: {
  title: string;
  description: string;
  accept: string;
  file: File | null;
  onChange: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <label className="fileCard">
      <span className="fileTitle">{title}</span>
      <span className="fileDescription">{description}</span>
      <input type="file" accept={accept} onChange={onChange} />
      <span className={file ? "fileName selected" : "fileName"}>
        {file ? file.name : "ファイルを選択"}
      </span>
    </label>
  );
}

async function readError(response: Response): Promise<string> {
  try {
    const data = await response.json();
    return typeof data.detail === "string" ? data.detail : "生成に失敗しました。";
  } catch {
    return "生成に失敗しました。";
  }
}

createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);

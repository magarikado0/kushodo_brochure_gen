import React, { ChangeEvent, FormEvent, useMemo, useRef, useState } from "react";
import { createRoot } from "react-dom/client";
import "./styles.css";

const API_BASE = import.meta.env.VITE_API_BASE ?? "";
const GITHUB_REPO_URL = "https://github.com/magarikado0/kushodo_brochure_gen";

function App() {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [formFile, setFormFile] = useState<File | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [message, setMessage] = useState<string>("");
  const [error, setError] = useState<string>("");

  const canSubmit = useMemo(() => Boolean(formFile && !isGenerating), [formFile, isGenerating]);

  const updateFormFile = (event: ChangeEvent<HTMLInputElement>) => {
    setFormFile(event.target.files?.[0] ?? null);
    setMessage("");
    setError("");
  };

  function resetFileSelection() {
    setFormFile(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  }

  async function submit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const selectedFile = fileInputRef.current?.files?.[0] ?? formFile;
    if (!selectedFile) {
      setError("作品情報フォームを選択してください。");
      return;
    }

    setIsGenerating(true);
    setMessage("生成しています。少しお待ちください。");
    setError("");

    try {
      const body = new FormData();
      body.append("form_file", selectedFile);

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
      resetFileSelection();
    } catch (caught) {
      if (caught instanceof TypeError) {
        setError(
          "通信に失敗しました。Excel でファイルを閉じてから、ファイルを選び直してもう一度お試しください。",
        );
      } else {
        setError(caught instanceof Error ? caught.message : "生成に失敗しました。");
      }
      setMessage("");
      resetFileSelection();
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <div className="layout">
      <main className="page">
        <header className="header">
          <h1>京大書道部パンフレット生成</h1>
          <p className="headerLead">
            作品情報フォーム（xlsx）をアップロードすると、作品一覧と作品紹介ページを含む Word ファイルを作成できます。
          </p>
        </header>

        <FileRequirements />

        <form className="card" onSubmit={submit}>
          <div className="cardHeader">
            <h2>ファイルを選択</h2>
            <p>上記の要件を満たした xlsx を選んでから、生成ボタンを押してください。</p>
          </div>

          <div className="uploadGrid">
            <FilePicker
              title="作品情報フォーム"
              description="Google フォームなどから出力した xlsx"
              accept=".xlsx"
              file={formFile}
              inputRef={fileInputRef}
              onChange={updateFormFile}
            />
          </div>

          <div className="actions">
            <button type="submit" className="submitButton" disabled={!canSubmit}>
              {isGenerating ? "生成中..." : "パンフレットを生成"}
            </button>
          </div>

          {message && <p className="feedback message">{message}</p>}
          {error && <ErrorFeedback message={error} />}
        </form>

        <aside className="noteCard">
          <h2>フィードバック</h2>
          <p>
            うまくいかないことや改善点・ご要望があれば、
            <a href={GITHUB_REPO_URL} target="_blank" rel="noopener noreferrer">
              GitHub リポジトリ
            </a>
            から Issue を書いたり、Pull Request を送ってください。
          </p>
        </aside>
      </main>
    </div>
  );
}

function FileRequirements() {
  return (
    <section className="requirementsCard">
      <details className="requirementsToggle">
        <summary className="requirementsSummary">
          <span className="requirementsSummaryMain">
            <span className="requirementsSummaryTitle">ファイルの要件</span>
            <span className="requirementsSummaryHint">形式が合わないと生成に失敗します</span>
          </span>
          <span className="requirementsSummaryAction" aria-hidden="true" />
        </summary>

        <div className="requirementsContent">
          <p className="requirementsIntro">アップロードする xlsx が、次の形式を満たしているか確認してください。</p>
          <article className="requirementPanel">
            <h3>作品情報フォーム（xlsx）</h3>
            <div className="requirementBody">
              <p>
                <code>個人</code> と <code>合作</code> の2シートが必要です。シート名にはそれぞれ「個人」「合作」を含めてください。
                列の順番は変わっても大丈夫ですが、列名の先頭部分は変えないでください。
              </p>
              <p>個人シートに必要な列:</p>
              <p className="requirementColumns">
                氏名、ふりがな、学年、臨書 or 創作、書体、作品名、作品の向き、作品サイズ、展示場所、表装形式、釈文、作品コメント
              </p>
              <p>合作シートに必要な列:</p>
              <p className="requirementColumns">
                合作参加者全員分、臨書 or 創作、書体、作品名、作品の向き、作品サイズ、展示場所、表装形式、釈文、作品コメント
              </p>
              <p>
                学年は <code>2回生</code>、<code>B2</code>、<code>修士1回生</code> などに対応しています。
                合作参加者は、名前の後ろに学年を付けてください（例: 加藤 杏次郎（B4） 星野 真帆（B4））。
              </p>
            </div>
          </article>
        </div>
      </details>
    </section>
  );
}

function FilePicker({
  title,
  description,
  accept,
  file,
  inputRef,
  onChange,
}: {
  title: string;
  description: string;
  accept: string;
  file: File | null;
  inputRef: React.RefObject<HTMLInputElement | null>;
  onChange: (event: ChangeEvent<HTMLInputElement>) => void;
}) {
  return (
    <label className={file ? "uploadBox uploadBoxSelected" : "uploadBox"}>
      <span className="uploadIcon" aria-hidden="true" />
      <span className="uploadTitle">{title}</span>
      <span className="uploadDescription">{description}</span>
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        onClick={(event) => {
          event.currentTarget.value = "";
        }}
        onChange={onChange}
      />
      <span className="uploadAction">{file ? file.name : "ファイルを選択"}</span>
    </label>
  );
}

function ErrorFeedback({ message }: { message: string }) {
  const lines = message.split("\n").map((line) => line.trimEnd()).filter((line) => line.length > 0);

  return (
    <div className="feedback error" role="alert">
      {lines.map((line, index) => (
        <p key={index} className={line.startsWith("直し方:") ? "errorHint" : undefined}>
          {line}
        </p>
      ))}
    </div>
  );
}

function formatApiErrorDetail(detail: unknown): string {
  if (typeof detail === "string") {
    return detail;
  }
  if (Array.isArray(detail)) {
    const messages = detail
      .map((item) => {
        if (typeof item === "string") {
          return item;
        }
        if (item && typeof item === "object" && "msg" in item) {
          return String(item.msg);
        }
        return null;
      })
      .filter((item): item is string => Boolean(item));
    return messages.length > 0 ? messages.join("\n") : "生成に失敗しました。";
  }
  return "生成に失敗しました。";
}

async function readError(response: Response): Promise<string> {
  try {
    const data = await response.json();
    return formatApiErrorDetail(data.detail);
  } catch {
    return "生成に失敗しました。";
  }
}

createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);

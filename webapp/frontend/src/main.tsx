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

  const canSubmit = useMemo(() => Boolean(formFile && !isGenerating), [formFile, isGenerating]);

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
    if (!formFile) {
      setError("作品情報フォームを選択してください。");
      return;
    }

    setIsGenerating(true);
    setMessage("生成しています。少しお待ちください。");
    setError("");

    try {
      const body = new FormData();
      body.append("form_file", formFile);
      if (templateFile) {
        body.append("template_file", templateFile);
      }

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
    <div className="layout">
      <main className="page">
        <header className="header">
          <h1>京大書道部パンフレット生成</h1>
          <p className="headerLead">
            作品情報フォームをアップロードすると、作品一覧と作品紹介ページを流し込んだ Word ファイルを生成します。
            パンフ鋳型を指定しない場合は、既定のパンフ鋳型を使います。
          </p>
        </header>

        <FileRequirements />

        <form className="card" onSubmit={submit}>
          <div className="cardHeader">
            <h2>ファイルを選択</h2>
            <p>上記の要件を満たしたファイルを選んでから、生成ボタンを押してください。パンフ鋳型は任意です。</p>
          </div>

          <div className="uploadGrid">
            <FilePicker
              title="作品情報フォーム"
              description="Google フォームなどから出力した xlsx"
              accept=".xlsx"
              file={formFile}
              onChange={updateFile("form")}
            />
            <FilePicker
              title="パンフ鋳型（任意）"
              description="前回パンフレットの docx を指定できます"
              hint="未選択のときは既定のパンフ鋳型を使用"
              accept=".docx"
              file={templateFile}
              onChange={updateFile("template")}
            />
          </div>

          <div className="actions">
            <button type="submit" className="submitButton" disabled={!canSubmit}>
              {isGenerating ? "生成中..." : "パンフレットを生成"}
            </button>
          </div>

          {message && <p className="feedback message">{message}</p>}
          {error && <p className="feedback error">{error}</p>}
        </form>

        <aside className="noteCard">
          <h2>生成後にすること</h2>
          <p>作品画像は自動では入りません。Word ファイル内の画像プレースホルダに手で配置してください。</p>
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
          <p className="requirementsIntro">アップロード前に、次の形式を満たしているか確認してください。</p>
          <div className="requirementsGrid">
            <article className="requirementPanel">
              <h3>作品情報フォーム（xlsx）</h3>
              <div className="requirementBody">
                <p>
                  <code>個人</code> と <code>合作</code> の2シートが必要です。シート名は変えないでください。
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

            <article className="requirementPanel">
              <h3>パンフ鋳型（docx）</h3>
              <div className="requirementBody">
                <p>
                  前回のパンフレットをコピーした Word ファイルを使ってください。完全に新しいファイルではなく、前回版をベースにしてください。
                  未指定の場合は、既定のパンフ鋳型が使われます。
                </p>
                <p>テンプレートには、次の内容が必要です。</p>
                <ul>
                  <li>見出し <code>作品一覧</code></li>
                  <li>見出し <code>個人作品</code></li>
                  <li>学年見出しの例として <code>一回生</code></li>
                  <li>作品紹介ページの開始位置として <code>賛助作品</code></li>
                  <li>作品紹介ページ内の、番号入りテキストボックス</li>
                  <li>後ろの固定ページの開始位置として <code>【顧問・部員紹介】</code></li>
                </ul>
              </div>
            </article>
          </div>
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
  onChange,
  hint,
}: {
  title: string;
  description: string;
  accept: string;
  file: File | null;
  onChange: (event: ChangeEvent<HTMLInputElement>) => void;
  hint?: string;
}) {
  return (
    <label className={file ? "uploadBox uploadBoxSelected" : "uploadBox"}>
      <span className="uploadIcon" aria-hidden="true" />
      <span className="uploadTitle">{title}</span>
      <span className="uploadDescription">{description}</span>
      <input type="file" accept={accept} onChange={onChange} />
      <span className="uploadAction">{file ? file.name : "ファイルを選択"}</span>
      {hint && !file && <span className="uploadHint">{hint}</span>}
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

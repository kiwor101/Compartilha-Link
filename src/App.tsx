import { useMemo, useState } from "react";
import { useMsal } from "@azure/msal-react";
import { AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { Check, Clipboard, CloudUpload, LogIn, LogOut, ShieldCheck, X } from "lucide-react";
import { getAccessToken, uploadFileAndCreateLink, type UploadedLink } from "./graph";

const maxSimpleUploadSize = 250 * 1024 * 1024;

type UploadStatus = "idle" | "uploading" | "done" | "error";

function App() {
  const { instance, accounts } = useMsal();
  const [sector, setSector] = useState("Geral");
  const [files, setFiles] = useState<File[]>([]);
  const [links, setLinks] = useState<UploadedLink[]>([]);
  const [status, setStatus] = useState<UploadStatus>("idle");
  const [message, setMessage] = useState("");
  const [copiedUrl, setCopiedUrl] = useState("");

  const account = accounts[0];
  const totalSize = useMemo(() => files.reduce((total, file) => total + file.size, 0), [files]);
  const hasOversizedFile = files.some((file) => file.size > maxSimpleUploadSize);

  async function signIn() {
    await instance.loginPopup({
      scopes: ["User.Read", "Files.ReadWrite"]
    });
  }

  async function signOut() {
    if (account) {
      await instance.logoutPopup({ account });
    }
  }

  async function uploadSelectedFiles() {
    if (!account || files.length === 0) {
      return;
    }

    if (hasOversizedFile) {
      setStatus("error");
      setMessage("Este modelo inicial aceita arquivos ate 250 MB. Arquivos maiores entram na proxima etapa.");
      return;
    }

    setStatus("uploading");
    setMessage("Enviando arquivos e criando links...");
    setLinks([]);

    try {
      const token = await getAccessToken(instance, account);
      const results: UploadedLink[] = [];

      for (const file of files) {
        const result = await uploadFileAndCreateLink(token, file, sector);
        results.push(result);
        setLinks([...results]);
      }

      setStatus("done");
      setMessage("Links criados com sucesso.");
    } catch (error) {
      setStatus("error");
      setMessage(error instanceof Error ? error.message : "Nao foi possivel concluir o envio.");
    }
  }

  async function copyLink(url: string) {
    await navigator.clipboard.writeText(url);
    setCopiedUrl(url);
    window.setTimeout(() => setCopiedUrl(""), 1800);
  }

  return (
    <main className="shell">
      <section className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Microsoft 365</p>
            <h1>Link Facil OneDrive</h1>
          </div>

          <AuthenticatedTemplate>
            <button className="ghostButton" onClick={signOut}>
              <LogOut size={18} />
              Sair
            </button>
          </AuthenticatedTemplate>
        </header>

        <UnauthenticatedTemplate>
          <section className="loginPanel">
            <div className="loginCopy">
              <ShieldCheck size={36} />
              <h2>Entre com sua conta corporativa</h2>
              <p>Depois do login, escolha os arquivos e receba links externos prontos para copiar.</p>
            </div>
            <button className="primaryButton" onClick={signIn}>
              <LogIn size={20} />
              Entrar com Microsoft
            </button>
          </section>
        </UnauthenticatedTemplate>

        <AuthenticatedTemplate>
          <section className="grid">
            <div className="panel">
              <div className="accountLine">
                <span>Conectado como</span>
                <strong>{account?.username}</strong>
              </div>

              <label className="field">
                <span>Setor ou categoria</span>
                <input
                  value={sector}
                  onChange={(event) => setSector(event.target.value)}
                  placeholder="Ex.: Prontuario, Prestacao de Contas"
                />
              </label>

              <label className="dropzone">
                <CloudUpload size={34} />
                <strong>Selecionar arquivos</strong>
                <span>PDFs, imagens e documentos ate 250 MB cada</span>
                <input
                  type="file"
                  multiple
                  onChange={(event) => setFiles(Array.from(event.target.files || []))}
                />
              </label>

              {files.length > 0 && (
                <div className="fileSummary">
                  <div>
                    <strong>{files.length} arquivo(s)</strong>
                    <span>{formatBytes(totalSize)} no total</span>
                  </div>
                  <button className="iconButton" onClick={() => setFiles([])} aria-label="Limpar arquivos">
                    <X size={18} />
                  </button>
                </div>
              )}

              <button
                className="primaryButton"
                disabled={files.length === 0 || status === "uploading"}
                onClick={uploadSelectedFiles}
              >
                <CloudUpload size={20} />
                {status === "uploading" ? "Enviando..." : "Enviar e gerar links"}
              </button>

              {message && <p className={`status ${status}`}>{message}</p>}
            </div>

            <div className="panel resultsPanel">
              <div className="panelHeader">
                <h2>Links gerados</h2>
                <span>{links.length}</span>
              </div>

              {links.length === 0 ? (
                <div className="emptyState">
                  <Clipboard size={30} />
                  <p>Os links aparecem aqui assim que o envio terminar.</p>
                </div>
              ) : (
                <div className="linkList">
                  {links.map((link) => (
                    <article className="linkItem" key={link.webUrl}>
                      <div>
                        <strong>{link.fileName}</strong>
                        <span>{formatBytes(link.size)}</span>
                      </div>
                      <button className="copyButton" onClick={() => copyLink(link.webUrl)}>
                        {copiedUrl === link.webUrl ? <Check size={18} /> : <Clipboard size={18} />}
                        {copiedUrl === link.webUrl ? "Copiado" : "Copiar"}
                      </button>
                    </article>
                  ))}
                </div>
              )}
            </div>
          </section>
        </AuthenticatedTemplate>
      </section>
    </main>
  );
}

function formatBytes(bytes: number) {
  if (bytes === 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;

  return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
}

export default App;

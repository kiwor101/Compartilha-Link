const graphBaseUrl = "https://graph.microsoft.com/v1.0";
const authScopes = "openid profile offline_access User.Read Files.ReadWrite";
const maxSimpleUploadSize = 250 * 1024 * 1024;
const uploadChunkSize = 10 * 1024 * 1024;
const historyFolderName = "_Compartilha Link Sistema";
const historyFileName = "historico.json";

const config = window.APP_CONFIG || {};
const authBaseUrl = `https://login.microsoftonline.com/${config.microsoftTenantId}/oauth2/v2.0`;
const redirectUri = buildRedirectUri();
const configIsReady =
  config.microsoftClientId &&
  config.microsoftTenantId &&
  !config.microsoftClientId.includes("cole-aqui") &&
  !config.microsoftTenantId.includes("cole-aqui");

let selectedFiles = [];
let account = null;
let tokenCache = readTokenCache();
let linkHistory = [];

const elements = {
  loginPanel: document.querySelector("#loginPanel"),
  appPanel: document.querySelector("#appPanel"),
  loginButton: document.querySelector("#loginButton"),
  logoutButton: document.querySelector("#logoutButton"),
  accountName: document.querySelector("#accountName"),
  sectorInput: document.querySelector("#sectorInput"),
  fileInput: document.querySelector("#fileInput"),
  fileSummary: document.querySelector("#fileSummary"),
  fileCount: document.querySelector("#fileCount"),
  fileSize: document.querySelector("#fileSize"),
  clearButton: document.querySelector("#clearButton"),
  uploadButton: document.querySelector("#uploadButton"),
  statusMessage: document.querySelector("#statusMessage"),
  linkCount: document.querySelector("#linkCount"),
  emptyState: document.querySelector("#emptyState"),
  linkList: document.querySelector("#linkList")
};

elements.loginButton.addEventListener("click", signIn);
elements.logoutButton.addEventListener("click", signOut);
elements.fileInput.addEventListener("change", () => {
  selectedFiles = Array.from(elements.fileInput.files || []);
  renderFileSummary();
});
elements.clearButton.addEventListener("click", clearFiles);
elements.uploadButton.addEventListener("click", uploadSelectedFiles);

initialize();

async function initialize() {
  if (!configIsReady) {
    elements.loginButton.disabled = true;
    setStatus("Configure o Client ID e o Tenant ID no arquivo config.js antes de entrar.", "error");
    return;
  }

  try {
    await finishRedirectLoginIfNeeded();
  } catch (error) {
    setStatus(error.message || "Nao foi possivel finalizar o login Microsoft.", "error");
  }

  if (tokenCache?.account) {
    account = tokenCache.account;
    renderSignedIn();
  }
}

async function signIn() {
  if (!configIsReady) {
    setStatus("Configure o Client ID e o Tenant ID no arquivo config.js antes de entrar.", "error");
    return;
  }

  const state = crypto.randomUUID();
  const verifier = base64UrlEncode(crypto.getRandomValues(new Uint8Array(32)));
  const challenge = await sha256Base64Url(verifier);

  sessionStorage.setItem("pkce_state", state);
  sessionStorage.setItem("pkce_verifier", verifier);

  const params = new URLSearchParams({
    client_id: config.microsoftClientId,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: authScopes,
    state,
    code_challenge: challenge,
    code_challenge_method: "S256"
  });

  window.location.href = `${authBaseUrl}/authorize?${params.toString()}`;
}

async function signOut() {
  clearTokenCache();
  renderSignedOut();

  const params = new URLSearchParams({
    post_logout_redirect_uri: redirectUri
  });
  window.location.href = `${authBaseUrl}/logout?${params.toString()}`;
}

function renderSignedIn() {
  elements.loginPanel.classList.add("hidden");
  elements.appPanel.classList.remove("hidden");
  elements.logoutButton.classList.remove("hidden");
  elements.accountName.textContent = account?.username || "-";
  loadAndRenderHistory();
}

function renderSignedOut() {
  elements.loginPanel.classList.remove("hidden");
  elements.appPanel.classList.add("hidden");
  elements.logoutButton.classList.add("hidden");
  elements.accountName.textContent = "-";
  renderFileSummary();
}

function renderFileSummary() {
  const totalSize = selectedFiles.reduce((total, file) => total + file.size, 0);
  elements.uploadButton.disabled = selectedFiles.length === 0;

  if (selectedFiles.length === 0) {
    elements.fileSummary.classList.add("hidden");
    return;
  }

  elements.fileSummary.classList.remove("hidden");
  elements.fileCount.textContent = `${selectedFiles.length} arquivo(s)`;
  elements.fileSize.textContent = `${formatBytes(totalSize)} no total`;
}

function clearFiles() {
  selectedFiles = [];
  elements.fileInput.value = "";
  renderFileSummary();
}

async function uploadSelectedFiles() {
  if (!account || selectedFiles.length === 0) {
    return;
  }

  elements.uploadButton.disabled = true;
  renderLinks([]);
  elements.emptyState.classList.add("hidden");
  elements.linkCount.textContent = "0";
  setStatus("Enviando arquivos e criando links...", "");

  try {
    const token = await getAccessToken();
    const results = [];
    const folderName = normalizeFolderName(elements.sectorInput.value);

    for (const file of selectedFiles) {
      setStatus(`Enviando ${file.name}...`, "");
      const result = await uploadFileAndCreateLink(token, file, folderName);
      results.push(result);
      renderLinks(results);
    }

    linkHistory = [...results, ...linkHistory];
    await saveHistory(token, linkHistory);
    renderLinks(linkHistory);
    setStatus("Links criados com sucesso.", "done");
  } catch (error) {
    setStatus(error.message || "Nao foi possivel concluir o envio.", "error");
  } finally {
    elements.uploadButton.disabled = selectedFiles.length === 0;
  }
}

async function getAccessToken() {
  tokenCache = readTokenCache();

  if (!tokenCache) {
    throw new Error("Faca login novamente para continuar.");
  }

  if (Date.now() < tokenCache.expiresAt - 60_000) {
    return tokenCache.accessToken;
  }

  if (!tokenCache.refreshToken) {
    throw new Error("Sua sessao expirou. Faca login novamente.");
  }

  const refreshed = await requestToken({
    grant_type: "refresh_token",
    client_id: config.microsoftClientId,
    scope: authScopes,
    refresh_token: tokenCache.refreshToken
  });

  saveTokenResponse(refreshed, tokenCache.account);
  return refreshed.access_token;
}

async function uploadFileAndCreateLink(accessToken, file, sector) {
  const rootFolder = config.uploadRootFolder || "Compartilhamentos Externos";
  const today = new Date().toISOString().slice(0, 10);
  const folderName = normalizeFolderName(sector);
  const folderPath = [rootFolder, folderName, today];

  await ensureFolderPath(accessToken, folderPath);

  const fileName = buildUniqueFileName(file.name);
  const uploadPath = [...folderPath, fileName].map(encodePathSegment).join("/");
  const uploadedItem =
    file.size <= maxSimpleUploadSize
      ? await uploadSmallFile(accessToken, uploadPath, file)
      : await uploadLargeFile(accessToken, uploadPath, file);

  const permission = await graphRequest(accessToken, `/me/drive/items/${uploadedItem.id}/createLink`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      type: "view",
      scope: "anonymous",
      retainInheritedPermissions: false
    })
  });

  if (!permission.link?.webUrl) {
    throw new Error("A Microsoft nao retornou um link compartilhavel para este arquivo.");
  }

  return {
    folderName,
    fileName,
    webUrl: permission.link.webUrl,
    size: uploadedItem.size,
    createdAt: new Date().toISOString()
  };
}

async function uploadSmallFile(accessToken, uploadPath, file) {
  return graphRequest(accessToken, `/me/drive/root:/${uploadPath}:/content`, {
    method: "PUT",
    headers: {
      "Content-Type": file.type || "application/octet-stream"
    },
    body: file
  });
}

async function uploadLargeFile(accessToken, uploadPath, file) {
  const session = await graphRequest(accessToken, `/me/drive/root:/${uploadPath}:/createUploadSession`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      item: {
        "@microsoft.graph.conflictBehavior": "replace"
      }
    })
  });

  let start = 0;
  let uploadedItem = null;

  while (start < file.size) {
    const end = Math.min(start + uploadChunkSize, file.size) - 1;
    const chunk = file.slice(start, end + 1);
    const response = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.size),
        "Content-Range": `bytes ${start}-${end}/${file.size}`
      },
      body: chunk
    });

    if (!response.ok) {
      throw await graphError(response);
    }

    const payload = await response.json();
    if (response.status === 201 || response.status === 200) {
      uploadedItem = payload;
    }

    start = end + 1;
    setStatus(`Enviando arquivo grande... ${Math.round((start / file.size) * 100)}%`, "");
  }

  if (!uploadedItem?.id) {
    throw new Error("A Microsoft nao confirmou a conclusao do upload grande.");
  }

  return uploadedItem;
}

async function ensureFolderPath(accessToken, folders) {
  const builtPath = [];

  for (const folder of folders) {
    const currentPath = [...builtPath, folder].map(encodePathSegment).join("/");
    const exists = await pathExists(accessToken, currentPath);

    if (!exists) {
      const parentPath = builtPath.map(encodePathSegment).join("/");
      const endpoint = parentPath
        ? `/me/drive/root:/${parentPath}:/children`
        : "/me/drive/root/children";

      await graphRequest(accessToken, endpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify({
          name: folder,
          folder: {},
          "@microsoft.graph.conflictBehavior": "fail"
        })
      });
    }

    builtPath.push(folder);
  }
}

async function pathExists(accessToken, path) {
  const response = await fetch(`${graphBaseUrl}/me/drive/root:/${path}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (response.status === 404) {
    return false;
  }

  if (!response.ok) {
    throw await graphError(response);
  }

  return true;
}

async function graphRequest(accessToken, endpoint, init) {
  const response = await fetch(`${graphBaseUrl}${endpoint}`, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(init.headers || {})
    }
  });

  if (!response.ok) {
    throw await graphError(response);
  }

  const history = await response.json();
  return Array.isArray(history) ? history : [];
}

async function graphError(response) {
  const payload = await response.json().catch(() => null);
  const message =
    payload?.error?.message ||
    `A Microsoft retornou erro ${response.status} ao processar a solicitacao.`;

  return new Error(message);
}

function renderLinks(links) {
  elements.linkCount.textContent = String(links.length);
  elements.linkList.innerHTML = "";

  if (links.length === 0) {
    elements.emptyState.classList.remove("hidden");
    return;
  }

  elements.emptyState.classList.add("hidden");

  for (const link of links) {
    const item = document.createElement("article");
    item.className = "linkItem";

    const info = document.createElement("div");
    const name = document.createElement("strong");
    const size = document.createElement("span");
    name.textContent = link.folderName || "Pasta sem nome";
    size.textContent = `${formatBytes(link.size)} - ${formatDate(link.createdAt)}`;
    info.append(name, size);

    const copyButton = document.createElement("button");
    copyButton.className = "copyButton";
    copyButton.type = "button";
    copyButton.textContent = "Copiar";
    copyButton.addEventListener("click", async () => {
      await navigator.clipboard.writeText(link.webUrl);
      copyButton.textContent = "Copiado";
      window.setTimeout(() => {
        copyButton.textContent = "Copiar";
      }, 1800);
    });

    item.append(info, copyButton);
    elements.linkList.append(item);
  }
}

async function loadAndRenderHistory() {
  try {
    const token = await getAccessToken();
    linkHistory = await loadHistory(token);
    renderLinks(linkHistory);
  } catch {
    linkHistory = [];
    renderLinks([]);
  }
}

async function loadHistory(accessToken) {
  const historyPath = getHistoryPath().map(encodePathSegment).join("/");
  const response = await fetch(`${graphBaseUrl}/me/drive/root:/${historyPath}:/content`, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (response.status === 404) {
    return [];
  }

  if (!response.ok) {
    throw await graphError(response);
  }

  return response.json();
}

async function saveHistory(accessToken, history) {
  const rootFolder = config.uploadRootFolder || "Compartilhamentos Externos";
  await ensureFolderPath(accessToken, [rootFolder, historyFolderName]);

  const historyPath = getHistoryPath().map(encodePathSegment).join("/");
  await graphRequest(accessToken, `/me/drive/root:/${historyPath}:/content`, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(history.slice(0, 200), null, 2)
  });
}

function getHistoryPath() {
  const rootFolder = config.uploadRootFolder || "Compartilhamentos Externos";
  return [rootFolder, historyFolderName, historyFileName];
}

function setStatus(message, type) {
  elements.statusMessage.textContent = message;
  elements.statusMessage.className = `status ${type}`;
}

function normalizeFolderName(value) {
  return value.trim() || "Geral";
}

function buildUniqueFileName(fileName) {
  const cleanName = fileName.replace(/[<>:"/\\|?*]+/g, "-").trim();
  const dotIndex = cleanName.lastIndexOf(".");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");

  if (dotIndex <= 0) {
    return `${cleanName || "arquivo"}-${stamp}`;
  }

  return `${cleanName.slice(0, dotIndex)}-${stamp}${cleanName.slice(dotIndex)}`;
}

function encodePathSegment(segment) {
  return encodeURIComponent(segment).replace(/%20/g, " ");
}

function formatBytes(bytes) {
  if (bytes === 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** index;

  return `${value.toFixed(value >= 10 || index === 0 ? 0 : 1)} ${units[index]}`;
}

function formatDate(value) {
  if (!value) {
    return "";
  }

  return new Intl.DateTimeFormat("pt-BR", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit"
  }).format(new Date(value));
}

async function finishRedirectLoginIfNeeded() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");
  const state = params.get("state");
  const error = params.get("error_description") || params.get("error");

  if (error) {
    window.history.replaceState({}, document.title, redirectUri);
    throw new Error(error);
  }

  if (!code) {
    return;
  }

  const expectedState = sessionStorage.getItem("pkce_state");
  const verifier = sessionStorage.getItem("pkce_verifier");

  if (!expectedState || !verifier || expectedState !== state) {
    throw new Error("A resposta de login nao bate com a solicitacao inicial.");
  }

  const tokenResponse = await requestToken({
    grant_type: "authorization_code",
    client_id: config.microsoftClientId,
    code,
    redirect_uri: redirectUri,
    scope: authScopes,
    code_verifier: verifier
  });

  const accountInfo = await getCurrentUser(tokenResponse.access_token);
  saveTokenResponse(tokenResponse, accountInfo);
  sessionStorage.removeItem("pkce_state");
  sessionStorage.removeItem("pkce_verifier");
  window.history.replaceState({}, document.title, redirectUri);
}

async function requestToken(fields) {
  const response = await fetch(`${authBaseUrl}/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded"
    },
    body: new URLSearchParams(fields)
  });

  if (!response.ok) {
    const payload = await response.json().catch(() => null);
    throw new Error(payload?.error_description || "A Microsoft recusou a solicitacao de token.");
  }

  return response.json();
}

async function getCurrentUser(accessToken) {
  const profile = await graphRequest(accessToken, "/me", {
    method: "GET"
  });

  return {
    username: profile.userPrincipalName || profile.mail || profile.displayName || "Usuario Microsoft"
  };
}

function saveTokenResponse(tokenResponse, accountInfo) {
  tokenCache = {
    accessToken: tokenResponse.access_token,
    refreshToken: tokenResponse.refresh_token || tokenCache?.refreshToken || null,
    expiresAt: Date.now() + Number(tokenResponse.expires_in || 3600) * 1000,
    account: accountInfo
  };

  localStorage.setItem("linkfacil_token", JSON.stringify(tokenCache));
}

function readTokenCache() {
  const raw = localStorage.getItem("linkfacil_token");
  if (!raw) {
    return null;
  }

  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
}

function clearTokenCache() {
  localStorage.removeItem("linkfacil_token");
  tokenCache = null;
  account = null;
  selectedFiles = [];
  elements.fileInput.value = "";
  elements.linkList.innerHTML = "";
  elements.linkCount.textContent = "0";
  elements.emptyState.classList.remove("hidden");
  linkHistory = [];
}

async function sha256Base64Url(value) {
  const bytes = new TextEncoder().encode(value);
  const hash = await crypto.subtle.digest("SHA-256", bytes);
  return base64UrlEncode(new Uint8Array(hash));
}

function base64UrlEncode(bytes) {
  let binary = "";
  for (const byte of bytes) {
    binary += String.fromCharCode(byte);
  }

  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
}

function buildRedirectUri() {
  const cleanPath = window.location.pathname.endsWith("/")
    ? window.location.pathname
    : window.location.pathname.replace(/\/[^/]*$/, "/");

  return `${window.location.origin}${cleanPath}`;
}

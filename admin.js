const graphBaseUrl = "https://graph.microsoft.com/v1.0";
const authScopes = "openid profile offline_access User.Read Files.ReadWrite";
const historyFolderName = "_Compartilha Link Sistema";
const historyFileName = "historico.json";
const allowedEmailDomain = "@santacasaandradina.org";

const config = window.APP_CONFIG || {};
const authBaseUrl = `https://login.microsoftonline.com/${config.microsoftTenantId}/oauth2/v2.0`;
const redirectUri = buildRedirectUri();
const rootFolder = config.uploadRootFolder || "Compartilhamentos Externos";

let account = null;
let tokenCache = readTokenCache();
let linkHistory = [];
let cleanupCandidates = [];

const elements = {
  loginPanel: document.querySelector("#loginPanel"),
  loginStatus: document.querySelector("#loginStatus"),
  adminPanel: document.querySelector("#adminPanel"),
  loginButton: document.querySelector("#loginButton"),
  logoutButton: document.querySelector("#logoutButton"),
  accountName: document.querySelector("#accountName"),
  ageSelect: document.querySelector("#ageSelect"),
  scanButton: document.querySelector("#scanButton"),
  deleteButton: document.querySelector("#deleteButton"),
  statusMessage: document.querySelector("#statusMessage"),
  resultCount: document.querySelector("#resultCount"),
  emptyState: document.querySelector("#emptyState"),
  resultList: document.querySelector("#resultList")
};

elements.loginButton.addEventListener("click", signIn);
elements.logoutButton.addEventListener("click", signOut);
elements.scanButton.addEventListener("click", scanOldFolders);
elements.deleteButton.addEventListener("click", deleteOldFolders);

initialize();

async function initialize() {
  try {
    await finishRedirectLoginIfNeeded();
  } catch (error) {
    setStatus(error.message || "Nao foi possivel finalizar o login Microsoft.", "error");
  }

  tokenCache = readTokenCache();
  if (tokenCache?.account) {
    if (!isCorporateAccount(tokenCache.account.username)) {
      clearTokenCache();
      renderSignedOut();
      setStatus(`Use uma conta corporativa do dominio ${allowedEmailDomain}.`, "error");
      return;
    }

    account = tokenCache.account;
    renderSignedIn();
  }
}

async function signIn() {
  const state = crypto.randomUUID();
  const verifier = base64UrlEncode(crypto.getRandomValues(new Uint8Array(32)));
  const challenge = await sha256Base64Url(verifier);

  sessionStorage.setItem("admin_pkce_state", state);
  sessionStorage.setItem("admin_pkce_verifier", verifier);

  const params = new URLSearchParams({
    client_id: config.microsoftClientId,
    response_type: "code",
    redirect_uri: redirectUri,
    response_mode: "query",
    scope: authScopes,
    state,
    code_challenge: challenge,
    code_challenge_method: "S256",
    prompt: "select_account"
  });

  window.location.href = `${authBaseUrl}/authorize?${params.toString()}`;
}

async function signOut() {
  clearTokenCache();
  sessionStorage.removeItem("admin_pkce_state");
  sessionStorage.removeItem("admin_pkce_verifier");
  window.history.replaceState({}, document.title, redirectUri);
  renderSignedOut();
  setStatus("Sessao encerrada nesta pagina.", "");
}

function renderSignedIn() {
  elements.loginPanel.classList.add("hidden");
  elements.adminPanel.classList.remove("hidden");
  elements.logoutButton.classList.remove("hidden");
  elements.accountName.textContent = account?.username || "-";
}

function renderSignedOut() {
  elements.loginPanel.classList.remove("hidden");
  elements.adminPanel.classList.add("hidden");
  elements.logoutButton.classList.add("hidden");
  elements.accountName.textContent = "-";
  renderCandidates([]);
}

async function scanOldFolders() {
  elements.scanButton.disabled = true;
  elements.deleteButton.disabled = true;
  setStatus("Lendo historico da conta logada...", "");

  try {
    const token = await getAccessToken();
    linkHistory = await loadHistory(token);
    cleanupCandidates = findExpiredHistoryItems(linkHistory, Number(elements.ageSelect.value));
    renderCandidates(cleanupCandidates);

    if (cleanupCandidates.length === 0) {
      setStatus("Nenhuma pasta com link vencido foi encontrada para esse periodo.", "done");
      return;
    }

    elements.deleteButton.disabled = false;
    setStatus(`${cleanupCandidates.length} pasta(s) encontrada(s). Confira a lista antes de apagar.`, "");
  } catch (error) {
    setStatus(error.message || "Nao foi possivel buscar as pastas antigas.", "error");
  } finally {
    elements.scanButton.disabled = false;
  }
}

async function deleteOldFolders() {
  if (cleanupCandidates.length === 0) {
    return;
  }

  const confirmed = window.confirm(
    `Voce esta prestes a apagar ${cleanupCandidates.length} pasta(s) dentro de ${rootFolder}. Esta acao nao pode ser desfeita pelo app. Deseja continuar?`
  );

  if (!confirmed) {
    return;
  }

  elements.scanButton.disabled = true;
  elements.deleteButton.disabled = true;
  setStatus("Apagando pastas antigas...", "");

  try {
    const token = await getAccessToken();
    const deletedKeys = new Set();

    for (const item of cleanupCandidates) {
      const folderPath = getFolderPath(item);
      setStatus(`Apagando ${item.folderName}...`, "");
      await deleteDriveItemByPath(token, folderPath);
      deletedKeys.add(getHistoryKey(item));
    }

    linkHistory = linkHistory.filter((item) => !deletedKeys.has(getHistoryKey(item)));
    await saveHistory(token, linkHistory);
    cleanupCandidates = [];
    renderCandidates([]);
    setStatus("Limpeza concluida e historico atualizado.", "done");
  } catch (error) {
    setStatus(error.message || "Nao foi possivel concluir a limpeza.", "error");
    elements.deleteButton.disabled = cleanupCandidates.length === 0;
  } finally {
    elements.scanButton.disabled = false;
  }
}

function findExpiredHistoryItems(history, years) {
  const threshold = subtractYears(new Date(), years);
  const seen = new Set();

  return (history || []).filter((item) => {
    const expiredAt = new Date(item.expiresAt || 0);
    const key = getHistoryKey(item);

    if (!item.folderName || Number.isNaN(expiredAt.getTime()) || expiredAt > threshold || seen.has(key)) {
      return false;
    }

    seen.add(key);
    return true;
  });
}

function renderCandidates(candidates) {
  elements.resultCount.textContent = String(candidates.length);
  elements.resultList.innerHTML = "";

  if (candidates.length === 0) {
    elements.emptyState.classList.remove("hidden");
    return;
  }

  elements.emptyState.classList.add("hidden");

  for (const item of candidates) {
    const article = document.createElement("article");
    article.className = "linkItem";

    const info = document.createElement("div");
    const name = document.createElement("strong");
    const details = document.createElement("span");
    const path = document.createElement("span");
    name.textContent = item.folderName || "Pasta sem nome";
    details.textContent = `${item.fileCount || item.files?.length || 0} arquivo(s) - vencido em ${formatDate(item.expiresAt)}`;
    path.textContent = getFolderPath(item);
    info.append(name, details, path);
    article.append(info);
    elements.resultList.append(article);
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
  const historyPath = getHistoryPath().map(encodePathSegment).join("/");
  await graphRequest(accessToken, `/me/drive/root:/${historyPath}:/content`, {
    method: "PUT",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(history.slice(0, 200), null, 2)
  });
}

async function deleteDriveItemByPath(accessToken, folderPath) {
  const response = await fetch(`${graphBaseUrl}/me/drive/root:/${folderPath}:`, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });

  if (response.status === 404) {
    return;
  }

  if (!response.ok) {
    throw await graphError(response);
  }
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

  return response.json();
}

async function graphError(response) {
  const payload = await response.json().catch(() => null);
  const message =
    payload?.error?.message ||
    `A Microsoft retornou erro ${response.status} ao processar a solicitacao.`;

  return new Error(message);
}

async function finishRedirectLoginIfNeeded() {
  const params = new URLSearchParams(window.location.search);
  const code = params.get("code");
  const state = params.get("state");
  const error = params.get("error_description") || params.get("error");

  if (error) {
    clearTokenCache();
    sessionStorage.removeItem("admin_pkce_state");
    sessionStorage.removeItem("admin_pkce_verifier");
    window.history.replaceState({}, document.title, redirectUri);
    throw new Error(error);
  }

  if (!code) {
    return;
  }

  const expectedState = sessionStorage.getItem("admin_pkce_state");
  const verifier = sessionStorage.getItem("admin_pkce_verifier");

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

  try {
    const accountInfo = await getCurrentUser(tokenResponse.access_token);
    saveTokenResponse(tokenResponse, accountInfo);
  } catch (error) {
    clearTokenCache();
    throw error;
  } finally {
    sessionStorage.removeItem("admin_pkce_state");
    sessionStorage.removeItem("admin_pkce_verifier");
    window.history.replaceState({}, document.title, redirectUri);
  }
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
  const username = profile.userPrincipalName || profile.mail || profile.displayName || "Usuario Microsoft";

  if (!isCorporateAccount(username)) {
    throw new Error(`Use uma conta corporativa do dominio ${allowedEmailDomain}.`);
  }

  return { username };
}

function saveTokenResponse(tokenResponse, accountInfo) {
  tokenCache = {
    accessToken: tokenResponse.access_token,
    refreshToken: tokenResponse.refresh_token || tokenCache?.refreshToken || null,
    expiresAt: Date.now() + Number(tokenResponse.expires_in || 3600) * 1000,
    account: accountInfo
  };

  localStorage.setItem("compartilha_admin_token", JSON.stringify(tokenCache));
}

function readTokenCache() {
  const raw = localStorage.getItem("compartilha_admin_token");
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
  localStorage.removeItem("compartilha_admin_token");
  tokenCache = null;
  account = null;
  cleanupCandidates = [];
}

function getHistoryPath() {
  return [rootFolder, historyFolderName, historyFileName];
}

function getFolderPath(item) {
  return (
    item.folderItemPath ||
    [rootFolder, sanitizePathSegment(item.folderName)]
      .map(encodePathSegment)
      .join("/")
  );
}

function getHistoryKey(item) {
  return `${item.folderName || ""}|${item.createdAt || ""}`;
}

function sanitizePathSegment(value) {
  return String(value || "").replace(/[<>:"/\\|?*]+/g, "-").trim() || "Geral";
}

function encodePathSegment(segment) {
  return encodeURIComponent(segment);
}

function subtractYears(date, years) {
  const result = new Date(date);
  result.setFullYear(result.getFullYear() - years);
  return result;
}

function isCorporateAccount(username) {
  return String(username || "").toLocaleLowerCase("pt-BR").endsWith(allowedEmailDomain);
}

function setStatus(message, type) {
  elements.statusMessage.textContent = message;
  elements.statusMessage.className = `status ${type}`;
  elements.loginStatus.textContent = message;
  elements.loginStatus.className = `status ${type}`;
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
  return `${window.location.origin}${window.location.pathname}`;
}

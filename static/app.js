const graphBaseUrl = "https://graph.microsoft.com/v1.0";
const authScopes = "openid profile offline_access User.Read Files.ReadWrite";
const maxSimpleUploadSize = 250 * 1024 * 1024;
const uploadChunkSize = 10 * 1024 * 1024;
const historyFolderName = "_Compartilha Link Sistema";
const historyFileName = "historico.json";
const allowedEmailDomain = "@santacasaandradina.org";
const defaultValidityDays = 90;

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
let uploadProgressState = null;

const elements = {
  loginPanel: document.querySelector("#loginPanel"),
  loginStatus: document.querySelector("#loginStatus"),
  appPanel: document.querySelector("#appPanel"),
  loginButton: document.querySelector("#loginButton"),
  logoutButton: document.querySelector("#logoutButton"),
  accountName: document.querySelector("#accountName"),
  sectorInput: document.querySelector("#sectorInput"),
  validitySelect: document.querySelector("#validitySelect"),
  dropzone: document.querySelector(".dropzone"),
  fileInput: document.querySelector("#fileInput"),
  fileSummary: document.querySelector("#fileSummary"),
  fileList: document.querySelector("#fileList"),
  fileCount: document.querySelector("#fileCount"),
  fileSize: document.querySelector("#fileSize"),
  clearButton: document.querySelector("#clearButton"),
  uploadProgress: document.querySelector("#uploadProgress"),
  progressPercent: document.querySelector("#progressPercent"),
  progressBar: document.querySelector("#progressBar"),
  uploadButton: document.querySelector("#uploadButton"),
  statusMessage: document.querySelector("#statusMessage"),
  linkCount: document.querySelector("#linkCount"),
  emptyState: document.querySelector("#emptyState"),
  historySearch: document.querySelector("#historySearch"),
  linkList: document.querySelector("#linkList")
};

elements.loginButton.addEventListener("click", signIn);
elements.logoutButton.addEventListener("click", signOut);
elements.fileInput.addEventListener("change", () => {
  selectedFiles = [...selectedFiles, ...mapSelectedFiles(elements.fileInput.files)];
  elements.fileInput.value = "";
  renderFileSummary();
});
elements.clearButton.addEventListener("click", clearFiles);
elements.uploadButton.addEventListener("click", uploadSelectedFiles);
elements.historySearch.addEventListener("input", () => renderLinks(linkHistory));
elements.dropzone.addEventListener("dragover", handleDragOver, true);
elements.dropzone.addEventListener("dragleave", handleDragLeave);
elements.dropzone.addEventListener("drop", handleDrop, true);

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
    code_challenge_method: "S256",
    prompt: "select_account"
  });

  window.location.href = `${authBaseUrl}/authorize?${params.toString()}`;
}

async function signOut() {
  clearTokenCache();
  sessionStorage.removeItem("pkce_state");
  sessionStorage.removeItem("pkce_verifier");
  window.history.replaceState({}, document.title, redirectUri);
  renderSignedOut();
  setStatus("Sessao encerrada neste aplicativo. Clique em Login para escolher a conta novamente.", "");
}

function renderSignedIn() {
  elements.loginPanel.classList.add("hidden");
  elements.appPanel.classList.remove("hidden");
  elements.logoutButton.classList.remove("hidden");
  elements.accountName.textContent = account?.username || "-";
  refreshSignedInUser();
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
  elements.fileList.innerHTML = "";

  if (selectedFiles.length === 0) {
    elements.fileSummary.classList.add("hidden");
    return;
  }

  elements.fileSummary.classList.remove("hidden");
  elements.fileCount.textContent = `${selectedFiles.length} arquivo(s)`;
  elements.fileSize.textContent = `${formatBytes(totalSize)} no total`;

  selectedFiles.forEach((file, index) => {
    const item = document.createElement("div");
    item.className = "fileItem";

    const info = document.createElement("div");
    const name = document.createElement("strong");
    const size = document.createElement("span");
    name.textContent = file.relativePath || file.name;
    size.textContent = formatBytes(file.size);
    info.append(name, size);

    const removeButton = document.createElement("button");
    removeButton.className = "removeFileButton";
    removeButton.type = "button";
    removeButton.textContent = "Remover";
    removeButton.addEventListener("click", () => {
      selectedFiles.splice(index, 1);
      renderFileSummary();
    });

    item.append(info, removeButton);
    elements.fileList.append(item);
  });
}

function clearFiles() {
  selectedFiles = [];
  elements.fileInput.value = "";
  renderFileSummary();
  hideUploadProgress();
}

async function uploadSelectedFiles() {
  if (!account || selectedFiles.length === 0) {
    return;
  }

  elements.uploadButton.disabled = true;
  initializeUploadProgress(selectedFiles);
  setStatus("Preparando envio...", "");

  try {
    const folderName = getRequiredFolderName();
    if (!folderName) {
      throw new Error("Informe o nome da pasta antes de enviar.");
    }

    const token = await getAccessToken();
    const results = [];
    const validityDays = getSelectedValidityDays();
    const expiresAt = buildExpirationDate(validityDays);

    linkHistory = await loadHistory(token);
    if (folderNameExistsInHistory(folderName, linkHistory)) {
      renderLinks(linkHistory);
      throw new Error("Ja existe um link gerado com esse nome de pasta. Altere o nome da pasta para continuar.");
    }

    const result = await uploadFilesAndCreateFolderLink(token, selectedFiles, folderName, {
      validityDays,
      expiresAt
    });
    results.push(result);

    linkHistory = [...results, ...linkHistory];
    await saveHistory(token, linkHistory);
    renderLinks(linkHistory);
    clearFiles();
    setStatus("Links criados com sucesso.", "done");
  } catch (error) {
    hideUploadProgress();
    handleAuthExpiredError(error);
    setStatus(getFriendlyErrorMessage(error, "Nao foi possivel concluir o envio."), "error");
  } finally {
    elements.uploadButton.disabled = selectedFiles.length === 0;
  }
}

function handleDragOver(event) {
  event.preventDefault();
  event.stopPropagation();
  elements.dropzone.classList.add("isDragging");
}

function handleDragLeave(event) {
  if (!elements.dropzone.contains(event.relatedTarget)) {
    elements.dropzone.classList.remove("isDragging");
  }
}

async function handleDrop(event) {
  event.preventDefault();
  event.stopPropagation();
  elements.dropzone.classList.remove("isDragging");
  setStatus("Lendo arquivos selecionados...", "");

  try {
    const droppedFiles = await readDroppedFiles(event.dataTransfer);

    if (droppedFiles.length === 0) {
      setStatus("Nenhum arquivo foi encontrado na selecao.", "error");
      return;
    }

    selectedFiles = [...selectedFiles, ...droppedFiles];
    renderFileSummary();
    setStatus("", "");
  } catch (error) {
    setStatus(error.message || "Nao foi possivel ler os arquivos arrastados.", "error");
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

  let refreshed;

  try {
    refreshed = await requestToken({
      grant_type: "refresh_token",
      client_id: config.microsoftClientId,
      scope: authScopes,
      refresh_token: tokenCache.refreshToken
    });
  } catch (error) {
    if (isAuthExpiredError(error)) {
      throw createSessionExpiredError();
    }

    throw error;
  }

  saveTokenResponse(refreshed, tokenCache.account);
  return refreshed.access_token;
}

async function uploadFilesAndCreateFolderLink(accessToken, files, sector, linkOptions = {}) {
  const rootFolder = config.uploadRootFolder || "Compartilhamentos Externos";
  const folderName = getNormalizedRequiredFolderName(sector);
  const folderPath = [rootFolder, folderName];
  const validityDays = linkOptions.validityDays || defaultValidityDays;
  const expiresAt = linkOptions.expiresAt || buildExpirationDate(validityDays);

  await ensureFolderPath(accessToken, folderPath);

  const uploadedFiles = [];

  for (const file of files) {
    const relativePath = normalizeRelativePath(file.relativePath || file.webkitRelativePath || file.name);
    const uploadParts = relativePath.split("/").filter(Boolean).map(sanitizePathSegment);

    if (uploadParts.length > 1) {
      await ensureFolderPath(accessToken, [...folderPath, ...uploadParts.slice(0, -1)]);
    }

    const uploadPath = [...folderPath, ...uploadParts].map(encodePathSegment).join("/");
    setStatus("Enviando arquivos...", "");
    const uploadedItem =
      file.size <= maxSimpleUploadSize
        ? await uploadSmallFile(accessToken, uploadPath, file)
        : await uploadLargeFile(accessToken, uploadPath, file, addUploadedBytes);

    if (file.size <= maxSimpleUploadSize) {
      addUploadedBytes(file.size);
    }

    setStatus("Criando link compartilhavel...", "");
    const filePermission = await createFileSharingLink(accessToken, uploadedItem.id, expiresAt);
    const fileWebUrl = extractSharingUrl(filePermission);

    if (!fileWebUrl) {
      throw new Error(
        `Arquivo enviado, mas a Microsoft nao retornou o link do arquivo ${relativePath}. Detalhe tecnico: ${JSON.stringify(
          filePermission
        ).slice(0, 1200)}`
      );
    }

    uploadedFiles.push({
      fileName: uploadParts.join("/"),
      id: uploadedItem.id,
      permissionId: filePermission.id || "",
      size: uploadedItem.size || file.size,
      webUrl: fileWebUrl
    });
  }

  const folderItemPath = folderPath.map(encodePathSegment).join("/");
  setStatus("Finalizando pasta e historico...", "");
  const folderItem = await graphRequest(accessToken, `/me/drive/root:/${folderItemPath}:`, {
    method: "GET"
  });

  let webUrl = "";
  let linkMode = "files";
  let fileLinks = uploadedFiles.map((file) => ({
    fileName: file.fileName,
    id: file.id,
    permissionId: file.permissionId,
    webUrl: file.webUrl,
    size: file.size
  }));
  let folderLinkError = "";
  let folderPermissionId = "";

  try {
    const permission = await createSharingLink(accessToken, folderItem.id, folderItemPath, expiresAt);
    webUrl = await resolveSharingUrl(accessToken, folderItem.id, permission);
    if (webUrl) {
      linkMode = "folder";
      fileLinks = [];
      folderPermissionId = permission.id || "";
    }
  } catch (error) {
    folderLinkError = error?.message || "Falha ao gerar link da pasta.";
    webUrl = "";
  }

  if (!webUrl) {
    webUrl = fileLinks[0]?.webUrl || "";
  }

  if (!webUrl) {
    throw new Error(
      `Nao foi possivel gerar link da pasta nem dos arquivos enviados. Detalhe tecnico: ${folderLinkError || "sem detalhe retornado"}`
    );
  }

  return {
    folderName,
    webUrl,
    linkMode,
    validityDays,
    expiresAt,
    folderItemId: folderItem.id,
    folderItemPath,
    folderPermissionId,
    fileLinks,
    size: uploadedFiles.reduce((total, file) => total + file.size, 0),
    fileCount: uploadedFiles.length,
    files: uploadedFiles,
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

async function uploadLargeFile(accessToken, uploadPath, file, onProgress) {
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
    onProgress?.(chunk.size);
  }

  if (!uploadedItem?.id) {
    throw new Error("A Microsoft nao confirmou a conclusao do upload grande.");
  }

  return uploadedItem;
}

async function createSharingLink(accessToken, itemId, folderItemPath, expiresAt) {
  const body = JSON.stringify({
    type: "view",
    scope: "anonymous",
    expirationDateTime: expiresAt
  });

  try {
    return await graphRequest(accessToken, `/me/drive/items/${itemId}/createLink`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body
    });
  } catch (firstError) {
    try {
      return await graphRequest(accessToken, `/me/drive/root:/${folderItemPath}:/createLink`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body
      });
    } catch {
      throw firstError;
    }
  }
}

async function createFileLinks(accessToken, uploadedFiles, expiresAt) {
  const links = [];

  for (const file of uploadedFiles) {
    if (!file.id) {
      continue;
    }

    try {
      const permission = await createFileSharingLink(accessToken, file.id, expiresAt);
      const webUrl = extractSharingUrl(permission) || (await resolveSharingUrl(accessToken, file.id, permission));

      if (!webUrl) {
        continue;
      }

      links.push({
        fileName: file.fileName,
        id: file.id,
        permissionId: permission.id || "",
        webUrl,
        size: file.size
      });
    } catch {
      continue;
    }
  }

  return links;
}

async function createFileSharingLink(accessToken, itemId, expiresAt) {
  return graphRequest(accessToken, `/me/drive/items/${itemId}/createLink`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      type: "view",
      scope: "anonymous",
      expirationDateTime: expiresAt,
      retainInheritedPermissions: false
    })
  });
}

async function resolveSharingUrl(accessToken, itemId, permission) {
  const directUrl = extractSharingUrl(permission);
  if (directUrl) {
    return directUrl;
  }

  const permissions = await graphRequest(accessToken, `/me/drive/items/${itemId}/permissions`, {
    method: "GET"
  });

  const sharedPermission = (permissions.value || []).find((item) => extractSharingUrl(item));
  return extractSharingUrl(sharedPermission);
}

function extractSharingUrl(permission) {
  if (!permission) {
    return "";
  }

  if (Array.isArray(permission.value)) {
    const item = permission.value.find((entry) => extractSharingUrl(entry));
    return extractSharingUrl(item);
  }

  return (
    permission?.link?.webUrl ||
    permission?.link?.weburl ||
    permission?.webUrl ||
    permission?.weburl ||
    permission?.shareLink?.webUrl ||
    permission?.shareLink?.weburl ||
    ""
  );
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

  return response.json();
}

async function graphError(response) {
  const payload = await response.json().catch(() => null);
  const message =
    payload?.error?.message ||
    `A Microsoft retornou erro ${response.status} ao processar a solicitacao.`;

  return new Error(message);
}

function renderLinks(links) {
  const filteredLinks = filterLinks(links);
  elements.linkCount.textContent = String(filteredLinks.length);
  elements.linkList.innerHTML = "";

  if (filteredLinks.length === 0) {
    elements.emptyState.classList.remove("hidden");
    const hasSearch = elements.historySearch.value.trim().length > 0;
    elements.emptyState.querySelector("p").textContent = hasSearch
      ? "Nenhum link encontrado para essa pesquisa."
      : "Os links aparecem aqui assim que o envio terminar.";
    return;
  }

  elements.emptyState.classList.add("hidden");

  for (const link of filteredLinks) {
    const item = document.createElement("article");
    item.className = "linkItem";

    const info = document.createElement("div");
    const name = document.createElement("strong");
    const size = document.createElement("span");
    name.textContent = link.folderName || "Pasta sem nome";
    const count = link.fileCount || link.files?.length || 1;
    size.textContent = `${count} arquivo(s) - ${formatBytes(link.size)} - criado em ${formatDate(link.createdAt)}`;
    const expires = document.createElement("span");
    const expired = isLinkExpired(link);
    expires.textContent = link.expiresAt
      ? `${expired ? "Link vencido em" : "Link ativo ate"} ${formatDate(link.expiresAt)}`
      : "Link sem validade registrada";
    info.append(name, size, expires);

    const actions = document.createElement("div");
    actions.className = "linkActions";
    const copyButton = document.createElement("button");
    copyButton.className = "copyButton";
    copyButton.type = "button";
    copyButton.textContent = link.linkMode === "files" ? "Copiar links" : "Copiar";
    copyButton.addEventListener("click", async () => {
      const clipboardText =
        link.linkMode === "files"
          ? (link.fileLinks || []).map((file) => `${file.fileName}: ${file.webUrl}`).join("\n")
          : link.webUrl;
      await navigator.clipboard.writeText(clipboardText);
      copyButton.textContent = "Copiado";
      window.setTimeout(() => {
        copyButton.textContent = link.linkMode === "files" ? "Copiar links" : "Copiar";
      }, 1800);
    });

    if (expired) {
      const renewSelect = document.createElement("select");
      renewSelect.className = "renewSelect";
      renewSelect.setAttribute("aria-label", "Prazo de renovacao");
      for (const days of [7, 30, 60, 90]) {
        const option = document.createElement("option");
        option.value = String(days);
        option.textContent = `${days} dias`;
        option.selected = Number(link.validityDays || defaultValidityDays) === days;
        renewSelect.append(option);
      }

      const renewButton = document.createElement("button");
      renewButton.className = "renewButton";
      renewButton.type = "button";
      renewButton.textContent = "Renovar";
      renewButton.addEventListener("click", () => renewLink(link, Number(renewSelect.value), renewButton));
      actions.append(renewSelect, renewButton);
    }

    actions.prepend(copyButton);
    item.append(info, actions);
    elements.linkList.append(item);
  }
}

async function renewLink(link, validityDays, button) {
  button.disabled = true;
  button.textContent = "Renovando...";
  setStatus("Renovando link...", "");

  try {
    const token = await getAccessToken();
    const renewedLink = await renewHistoryLink(token, link, validityDays);
    linkHistory = linkHistory.map((item) => (isSameHistoryItem(item, link) ? renewedLink : item));
    await saveHistory(token, linkHistory);
    renderLinks(linkHistory);
    setStatus("Link renovado com sucesso.", "done");
  } catch (error) {
    handleAuthExpiredError(error);
    setStatus(getFriendlyErrorMessage(error, "Nao foi possivel renovar o link."), "error");
    button.disabled = false;
    button.textContent = "Renovar";
  }
}

async function renewHistoryLink(accessToken, link, validityDays) {
  const expiresAt = buildExpirationDate(validityDays);
  const folderItemPath =
    link.folderItemPath ||
    [config.uploadRootFolder || "Compartilhamentos Externos", normalizeFolderName(link.folderName)]
      .map(encodePathSegment)
      .join("/");
  const folderItem = link.folderItemId
    ? { id: link.folderItemId }
    : await getDriveItemByPath(accessToken, folderItemPath);

  if (link.linkMode !== "files") {
    const folderPermissionId = link.folderPermissionId || (await findPermissionIdByUrl(accessToken, folderItem.id, link.webUrl));

    if (folderPermissionId) {
      await deletePermission(accessToken, folderItem.id, folderPermissionId);
    }

    const permission = await createSharingLink(accessToken, folderItem.id, folderItemPath, expiresAt);
    const webUrl = await resolveSharingUrl(accessToken, folderItem.id, permission);

    if (!webUrl) {
      throw new Error("A Microsoft nao retornou o novo link da pasta.");
    }

    return {
      ...link,
      webUrl,
      linkMode: "folder",
      validityDays,
      expiresAt,
      folderItemId: folderItem.id,
      folderItemPath,
      folderPermissionId: permission.id || "",
      renewedAt: new Date().toISOString()
    };
  }

  const renewedFiles = [];

  for (const file of link.fileLinks || link.files || []) {
    const relativeFilePath = normalizeRelativePath(file.fileName).split("/").map(encodePathSegment).join("/");
    const fileItem = file.id ? { id: file.id } : await getDriveItemByPath(accessToken, `${folderItemPath}/${relativeFilePath}`);

    const filePermissionId = file.permissionId || (await findPermissionIdByUrl(accessToken, fileItem.id, file.webUrl));

    if (filePermissionId) {
      await deletePermission(accessToken, fileItem.id, filePermissionId);
    }

    const permission = await createFileSharingLink(accessToken, fileItem.id, expiresAt);
    const webUrl = extractSharingUrl(permission) || (await resolveSharingUrl(accessToken, fileItem.id, permission));

    if (!webUrl) {
      throw new Error(`A Microsoft nao retornou o novo link do arquivo ${file.fileName}.`);
    }

    renewedFiles.push({
      ...file,
      id: fileItem.id,
      permissionId: permission.id || "",
      webUrl
    });
  }

  return {
    ...link,
    webUrl: renewedFiles[0]?.webUrl || "",
    linkMode: "files",
    validityDays,
    expiresAt,
    fileLinks: renewedFiles,
    renewedAt: new Date().toISOString()
  };
}

async function getDriveItemByPath(accessToken, itemPath) {
  return graphRequest(accessToken, `/me/drive/root:/${itemPath}:`, {
    method: "GET"
  });
}

async function deletePermission(accessToken, itemId, permissionId) {
  const response = await fetch(`${graphBaseUrl}/me/drive/items/${itemId}/permissions/${permissionId}`, {
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

async function findPermissionIdByUrl(accessToken, itemId, webUrl) {
  if (!webUrl) {
    return "";
  }

  const permissions = await graphRequest(accessToken, `/me/drive/items/${itemId}/permissions`, {
    method: "GET"
  });
  const matchingPermission = (permissions.value || []).find((permission) => extractSharingUrl(permission) === webUrl);
  return matchingPermission?.id || "";
}

function isSameHistoryItem(item, target) {
  return item.createdAt === target.createdAt && item.folderName === target.folderName;
}

function isLinkExpired(link) {
  return Boolean(link.expiresAt) && new Date(link.expiresAt).getTime() <= Date.now();
}

async function refreshSignedInUser() {
  try {
    const token = await getAccessToken();
    const accountInfo = await getCurrentUser(token);
    account = accountInfo;
    elements.accountName.textContent = accountInfo.username;
    if (tokenCache) {
      saveTokenResponse(
        {
          access_token: tokenCache.accessToken,
          refresh_token: tokenCache.refreshToken,
          expires_in: Math.max(60, Math.floor((tokenCache.expiresAt - Date.now()) / 1000))
        },
        accountInfo
      );
    }
  } catch (error) {
    handleAuthExpiredError(error);
    elements.accountName.textContent = account?.username || "Usuario Microsoft";
  }
}

async function loadAndRenderHistory() {
  try {
    const token = await getAccessToken();
    linkHistory = await loadHistory(token);
    renderLinks(linkHistory);
  } catch (error) {
    handleAuthExpiredError(error);
    linkHistory = [];
    renderLinks([]);
    if (isAuthExpiredError(error)) {
      setStatus(getFriendlyErrorMessage(error, "Sua sessao expirou."), "error");
    }
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
  elements.loginStatus.textContent = message;
  elements.loginStatus.className = `status ${type}`;
}

function initializeUploadProgress(files) {
  const totalBytes = files.reduce((total, file) => total + file.size, 0);
  uploadProgressState = {
    totalBytes: Math.max(totalBytes, 1),
    uploadedBytes: 0
  };
  elements.uploadProgress.classList.remove("hidden");
  updateUploadProgress();
}

function addUploadedBytes(bytes) {
  if (!uploadProgressState) {
    return;
  }

  uploadProgressState.uploadedBytes = Math.min(
    uploadProgressState.totalBytes,
    uploadProgressState.uploadedBytes + Math.max(bytes, 0)
  );
  updateUploadProgress();
}

function updateUploadProgress() {
  if (!uploadProgressState) {
    return;
  }

  const percent = Math.min(
    100,
    Math.round((uploadProgressState.uploadedBytes / uploadProgressState.totalBytes) * 100)
  );
  elements.progressPercent.textContent = `${percent}%`;
  elements.progressBar.style.width = `${percent}%`;
}

function hideUploadProgress() {
  uploadProgressState = null;
  elements.uploadProgress.classList.add("hidden");
  elements.progressPercent.textContent = "0%";
  elements.progressBar.style.width = "0%";
}

function normalizeFolderName(value) {
  return sanitizePathSegment(value.trim() || "Geral");
}

function getRequiredFolderName() {
  return getNormalizedRequiredFolderName(elements.sectorInput.value);
}

function getNormalizedRequiredFolderName(value) {
  const rawValue = String(value || "").trim();
  return rawValue ? sanitizePathSegment(rawValue) : "";
}

function sanitizePathSegment(value) {
  return value.replace(/[<>:"/\\|?*]+/g, "-").trim() || "arquivo";
}

function normalizeRelativePath(value) {
  return value
    .split("/")
    .filter(Boolean)
    .map(sanitizePathSegment)
    .join("/");
}

function folderNameExistsInHistory(folderName, history) {
  const target = normalizeHistoryName(folderName);
  return (history || []).some((item) => normalizeHistoryName(item.folderName) === target);
}

function filterLinks(links) {
  const search = normalizeSearchText(elements.historySearch.value);

  if (!search) {
    return links || [];
  }

  return (links || []).filter((link) => {
    const searchableText = [
      link.folderName,
      formatDate(link.createdAt),
      formatDate(link.expiresAt),
      link.createdAt,
      link.expiresAt
    ]
      .filter(Boolean)
      .map(normalizeSearchText)
      .join(" ");

    return searchableText.includes(search);
  });
}

function normalizeSearchText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("pt-BR")
    .trim();
}

function normalizeHistoryName(value) {
  return normalizeFolderName(String(value || "")).toLocaleLowerCase("pt-BR");
}

function mapSelectedFiles(fileList) {
  return Array.from(fileList || []).map((file) => {
    file.relativePath = file.webkitRelativePath || file.name;
    return file;
  });
}

async function readDroppedFiles(dataTransfer) {
  const transferFiles = mapSelectedFiles(dataTransfer?.files || []);
  const items = Array.from(dataTransfer?.items || []);
  const supportsDirectoryDrop = items.some((item) => typeof item.webkitGetAsEntry === "function");

  if (!supportsDirectoryDrop) {
    return transferFiles;
  }

  const files = [];

  for (const item of items) {
    const entry = item.webkitGetAsEntry();

    if (!entry) {
      const file = item.getAsFile();
      if (file) {
        file.relativePath = file.name;
        files.push(file);
      }
      continue;
    }

    files.push(...(await readEntryFiles(entry, "")));
  }

  return mergeUniqueFiles(files, transferFiles);
}

function mergeUniqueFiles(primaryFiles, fallbackFiles) {
  const merged = [];
  const seen = new Set();

  for (const file of [...primaryFiles, ...fallbackFiles]) {
    const key = `${file.relativePath || file.webkitRelativePath || file.name}|${file.size}|${file.lastModified}`;
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    file.relativePath = file.relativePath || file.webkitRelativePath || file.name;
    merged.push(file);
  }

  return merged;
}

async function readEntryFiles(entry, parentPath) {
  if (entry.isFile) {
    const file = await getFileFromEntry(entry);
    file.relativePath = `${parentPath}${file.name}`;
    return [file];
  }

  if (!entry.isDirectory) {
    return [];
  }

  const directoryPath = `${parentPath}${entry.name}/`;
  const reader = entry.createReader();
  const entries = await readAllDirectoryEntries(reader);
  const files = await Promise.all(entries.map((child) => readEntryFiles(child, directoryPath)));
  return files.flat();
}

function getFileFromEntry(entry) {
  return new Promise((resolve, reject) => {
    entry.file(resolve, reject);
  });
}

function readAllDirectoryEntries(reader) {
  return new Promise((resolve, reject) => {
    const entries = [];

    function readBatch() {
      reader.readEntries((batch) => {
        if (batch.length === 0) {
          resolve(entries);
          return;
        }

        entries.push(...batch);
        readBatch();
      }, reject);
    }

    readBatch();
  });
}

function encodePathSegment(segment) {
  return encodeURIComponent(segment);
}

function getSelectedValidityDays() {
  const days = Number(elements.validitySelect.value);
  return [7, 30, 60, 90].includes(days) ? days : defaultValidityDays;
}

function buildExpirationDate(days) {
  const expiresAt = new Date(Date.now() + Number(days) * 24 * 60 * 60 * 1000);
  return formatGraphDateTime(expiresAt);
}

function formatGraphDateTime(date) {
  return date.toISOString().replace(/\.\d{3}Z$/, "Z");
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
    clearTokenCache();
    sessionStorage.removeItem("pkce_state");
    sessionStorage.removeItem("pkce_verifier");
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

  try {
    const accountInfo = await getCurrentUser(tokenResponse.access_token);
    saveTokenResponse(tokenResponse, accountInfo);
  } catch (error) {
    clearTokenCache();
    throw error;
  } finally {
    sessionStorage.removeItem("pkce_state");
    sessionStorage.removeItem("pkce_verifier");
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
    const error = new Error(payload?.error_description || "A Microsoft recusou a solicitacao de token.");
    error.code = payload?.error || "";
    error.description = payload?.error_description || "";
    throw error;
  }

  return response.json();
}

function createSessionExpiredError() {
  const error = new Error(
    "Sua sessao Microsoft expirou. Clique em Login novamente e selecione os arquivos para reenviar."
  );
  error.name = "SessionExpiredError";
  return error;
}

function isAuthExpiredError(error) {
  const text = `${error?.code || ""} ${error?.description || ""} ${error?.message || ""}`;
  return (
    error?.name === "SessionExpiredError" ||
    text.includes("AADSTS700084") ||
    text.includes("refresh token") ||
    text.includes("invalid_grant")
  );
}

function handleAuthExpiredError(error) {
  if (!isAuthExpiredError(error)) {
    return;
  }

  clearTokenCache();
  sessionStorage.removeItem("pkce_state");
  sessionStorage.removeItem("pkce_verifier");
  renderSignedOut();
}

function getFriendlyErrorMessage(error, fallbackMessage) {
  if (isAuthExpiredError(error)) {
    return createSessionExpiredError().message;
  }

  return error?.message || fallbackMessage;
}

async function getCurrentUser(accessToken) {
  const profile = await graphRequest(accessToken, "/me", {
    method: "GET"
  });
  const username = profile.userPrincipalName || profile.mail || profile.displayName || "Usuario Microsoft";
  validateCorporateAccount(username);

  return {
    username
  };
}

function validateCorporateAccount(username) {
  if (!isCorporateAccount(username)) {
    throw new Error(`Use uma conta corporativa do dominio ${allowedEmailDomain}.`);
  }
}

function isCorporateAccount(username) {
  return String(username || "").toLocaleLowerCase("pt-BR").endsWith(allowedEmailDomain);
}

function saveTokenResponse(tokenResponse, accountInfo) {
  const tokenAccount = accountInfo?.username ? accountInfo : getAccountFromIdToken(tokenResponse.id_token);

  tokenCache = {
    accessToken: tokenResponse.access_token,
    refreshToken: tokenResponse.refresh_token || tokenCache?.refreshToken || null,
    expiresAt: Date.now() + Number(tokenResponse.expires_in || 3600) * 1000,
    account: tokenAccount
  };

  localStorage.setItem("linkfacil_token", JSON.stringify(tokenCache));
}

function getAccountFromIdToken(idToken) {
  if (!idToken) {
    return { username: "Usuario Microsoft" };
  }

  try {
    const payload = JSON.parse(atob(idToken.split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
    return {
      username: payload.preferred_username || payload.upn || payload.email || payload.name || "Usuario Microsoft"
    };
  } catch {
    return { username: "Usuario Microsoft" };
  }
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

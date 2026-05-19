import type { AccountInfo, IPublicClientApplication } from "@azure/msal-browser";

const graphBaseUrl = "https://graph.microsoft.com/v1.0";
const scopes = ["User.Read", "Files.ReadWrite"];

export type UploadedLink = {
  fileName: string;
  webUrl: string;
  size: number;
};

export async function getAccessToken(
  instance: IPublicClientApplication,
  account: AccountInfo
) {
  const request = { scopes, account };

  try {
    const response = await instance.acquireTokenSilent(request);
    return response.accessToken;
  } catch {
    const response = await instance.acquireTokenPopup(request);
    return response.accessToken;
  }
}

export async function uploadFileAndCreateLink(
  accessToken: string,
  file: File,
  sector: string
): Promise<UploadedLink> {
  const rootFolder = import.meta.env.VITE_UPLOAD_ROOT_FOLDER || "Compartilhamentos Externos";
  const today = new Date().toISOString().slice(0, 10);
  const folderPath = [rootFolder, normalizeFolderName(sector), today];

  await ensureFolderPath(accessToken, folderPath);

  const uniqueFileName = buildUniqueFileName(file.name);
  const uploadPath = [...folderPath, uniqueFileName].map(encodePathSegment).join("/");
  const uploadedItem = await graphRequest<{ id: string; size: number }>(
    accessToken,
    `/me/drive/root:/${uploadPath}:/content`,
    {
      method: "PUT",
      headers: {
        "Content-Type": file.type || "application/octet-stream"
      },
      body: file
    }
  );

  const permission = await graphRequest<{ link?: { webUrl?: string } }>(
    accessToken,
    `/me/drive/items/${uploadedItem.id}/createLink`,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify({
        type: "view",
        scope: "anonymous",
        retainInheritedPermissions: false
      })
    }
  );

  if (!permission.link?.webUrl) {
    throw new Error("A Microsoft nao retornou um link compartilhavel para este arquivo.");
  }

  return {
    fileName: uniqueFileName,
    webUrl: permission.link.webUrl,
    size: uploadedItem.size
  };
}

async function ensureFolderPath(accessToken: string, folders: string[]) {
  const builtPath: string[] = [];

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

async function pathExists(accessToken: string, path: string) {
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

async function graphRequest<T>(
  accessToken: string,
  endpoint: string,
  init: RequestInit
): Promise<T> {
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

  return response.json() as Promise<T>;
}

async function graphError(response: Response) {
  const payload = await response.json().catch(() => null);
  const message =
    payload?.error?.message ||
    `A Microsoft retornou erro ${response.status} ao processar a solicitacao.`;

  return new Error(message);
}

function normalizeFolderName(value: string) {
  return value.trim() || "Geral";
}

function buildUniqueFileName(fileName: string) {
  const cleanName = fileName.replace(/[<>:"/\\|?*]+/g, "-").trim();
  const dotIndex = cleanName.lastIndexOf(".");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");

  if (dotIndex <= 0) {
    return `${cleanName || "arquivo"}-${stamp}`;
  }

  return `${cleanName.slice(0, dotIndex)}-${stamp}${cleanName.slice(dotIndex)}`;
}

function encodePathSegment(segment: string) {
  return encodeURIComponent(segment).replace(/%20/g, " ");
}

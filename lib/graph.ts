import { createHmac, timingSafeEqual } from "node:crypto";

export class PortalError extends Error {
  status: number;
  constructor(status: number, message: string) {
    super(message);
    this.status = status;
  }
}
export type Item = {
  id: string;
  name: string;
  size?: number;
  folder?: { childCount: number };
  file?: { mimeType: string };
  parentReference?: { id?: string; driveId?: string };
  remoteItem?: unknown;
};
function required(name: string) {
  const value = process.env[name];
  if (!value)
    throw new PortalError(503, "Het portaal is nog niet geconfigureerd.");
  return value;
}
function config() {
  return {
    drive: required("ONEDRIVE_DRIVE_ID"),
    root: required("ONEDRIVE_ROOT_FOLDER_ID"),
  };
}
const enc = encodeURIComponent;
let token: { value: string; until: number } | undefined;
async function accessToken() {
  if (token && Date.now() < token.until) return token.value;
  const response = await fetch(
    `https://login.microsoftonline.com/${enc(required("AZURE_TENANT_ID"))}/oauth2/v2.0/token`,
    {
      method: "POST",
      cache: "no-store",
      signal: AbortSignal.timeout(15000),
      body: new URLSearchParams({
        client_id: required("AZURE_CLIENT_ID"),
        client_secret: required("AZURE_CLIENT_SECRET"),
        grant_type: "client_credentials",
        scope: "https://graph.microsoft.com/.default",
      }),
    },
  );
  if (!response.ok)
    throw new PortalError(
      503,
      "Verbinding met Microsoft 365 is tijdelijk niet beschikbaar.",
    );
  const data = await response.json();
  token = {
    value: data.access_token,
    until: Date.now() + Math.max(0, data.expires_in - 60) * 1000,
  };
  return token.value;
}
const graphBase = "https://graph.microsoft.com/v1.0";
async function graphResponse(path: string) {
  const url = new URL(path.startsWith("https:") ? path : graphBase + path);
  if (
    url.origin !== "https://graph.microsoft.com" ||
    !url.pathname.startsWith("/v1.0/")
  )
    throw new PortalError(502, "Ongeldig antwoord van Microsoft 365.");
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${await accessToken()}` },
    cache: "no-store",
    redirect: "manual",
    signal: AbortSignal.timeout(20000),
  });
  if (response.status === 404)
    throw new PortalError(404, "Deze map of dit bestand bestaat niet meer.");
  if (!response.ok && response.status !== 302)
    throw new PortalError(
      503,
      "Microsoft 365 is tijdelijk niet beschikbaar. Probeer het later opnieuw.",
    );
  return response;
}
function itemPath(id: string) {
  return `/drives/${enc(config().drive)}/items/${enc(id)}`;
}
export async function getItem(id: string): Promise<Item> {
  return (
    await graphResponse(
      itemPath(id) +
        "?$select=id,name,size,folder,file,parentReference,remoteItem",
    )
  ).json();
}
export async function children(id: string): Promise<Item[]> {
  let next: string | undefined =
    itemPath(id) +
    "/children?$select=id,name,size,folder,file,parentReference,remoteItem&$top=200";
  const result: Item[] = [];
  while (next) {
    const data = await (await graphResponse(next)).json();
    result.push(...data.value.filter((item: Item) => !item.remoteItem));
    next = data["@odata.nextLink"];
  }
  return result;
}
export function projectId(folderId: string) {
  const secret = required("PROJECT_LINK_SECRET");
  if (Buffer.byteLength(secret) < 32)
    throw new PortalError(
      503,
      "De projectlinksleutel moet minstens 32 bytes bevatten.",
    );
  return createHmac("sha256", secret)
    .update(`${config().drive}:${folderId}`)
    .digest("base64url");
}
export async function projects() {
  return (await children(config().root))
    .filter((item) => item.folder)
    .map((item) => ({ ...item, projectId: projectId(item.id) }));
}
export async function resolveProject(id: string) {
  if (!/^[A-Za-z0-9_-]{43}$/.test(id))
    throw new PortalError(
      404,
      "Deze projectlink is ongeldig of niet meer beschikbaar.",
    );
  const project = (await projects()).find((item) =>
    timingSafeEqual(Buffer.from(id), Buffer.from(item.projectId)),
  );
  if (!project)
    throw new PortalError(
      404,
      "Deze projectlink is ongeldig of niet meer beschikbaar.",
    );
  return project;
}
// Never trust a client-supplied item ID: follow its real ancestry in the configured drive.
export async function authorizedItem(project: Item, id: string): Promise<Item> {
  const original = await getItem(id);
  let current = original;
  const visited = new Set<string>();
  for (let depth = 0; depth < 200; depth++) {
    if (
      current.remoteItem ||
      (current.parentReference?.driveId &&
        current.parentReference.driveId !== config().drive)
    )
      break;
    if (current.id === project.id) return original;
    if (
      visited.has(current.id) ||
      !current.parentReference?.id ||
      current.id === config().root
    )
      break;
    visited.add(current.id);
    current = await getItem(current.parentReference.id);
  }
  throw new PortalError(
    404,
    "Deze map of dit bestand hoort niet bij dit project.",
  );
}
export function previewType(item: Item) {
  const mime = item.file?.mimeType;
  return mime &&
    [
      "application/pdf",
      "image/jpeg",
      "image/png",
      "image/gif",
      "image/webp",
      "image/avif",
    ].includes(mime)
    ? mime
    : undefined;
}
export async function fileContent(item: Item, range: string | null) {
  const response = await graphResponse(itemPath(item.id) + "/content");
  if (response.status !== 302) return response;
  const location = response.headers.get("location");
  if (!location || new URL(location).protocol !== "https:")
    throw new PortalError(502, "Download tijdelijk niet beschikbaar.");
  // Preauthenticated URL stays on the server; never forward the Graph Authorization header.
  const content = await fetch(location, {
    cache: "no-store",
    headers: range ? { Range: range } : {},
    signal: AbortSignal.timeout(30000),
  });
  if (!content.ok && content.status !== 416)
    throw new PortalError(503, "Download tijdelijk niet beschikbaar.");
  return content;
}
export function publicError(error: unknown) {
  return error instanceof PortalError
    ? error
    : new PortalError(
        503,
        "Het portaal is tijdelijk niet beschikbaar. Probeer het later opnieuw.",
      );
}

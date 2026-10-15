import { DEMO_PROJECT_ID, demoItem, demoContent } from "@/lib/demo";
import {
  authorizedItem,
  fileContent,
  PortalError,
  previewType,
  publicError,
  resolveProject,
} from "@/lib/graph";

export const runtime = "nodejs";
export async function GET(
  request: Request,
  context: { params: Promise<{ projectId: string; itemId: string }> },
) {
  const headers = new Headers({
    "Cache-Control": "private, no-store",
    "X-Content-Type-Options": "nosniff",
    "Referrer-Policy": "no-referrer",
    "Content-Security-Policy":
      "sandbox; default-src 'none'; frame-ancestors 'self'",
  });
  try {
    const { projectId, itemId } = await context.params;
    const isDemo = projectId === DEMO_PROJECT_ID;
    const item = isDemo
      ? demoItem(itemId)
      : await authorizedItem(await resolveProject(projectId), itemId);
    if (!item.file) throw new PortalError(404, "Bestand niet gevonden.");
    const inline =
      new URL(request.url).searchParams.get("preview") === "1" &&
      previewType(item);
    const upstream = isDemo
      ? await demoContent(itemId)
      : await fileContent(item, request.headers.get("range"));
    headers.set("Content-Type", inline || "application/octet-stream");
    headers.set(
      "Content-Disposition",
      `${inline ? "inline" : "attachment"}; filename="document"; filename*=UTF-8''${encodeURIComponent(item.name).replace(/['()*]/g, (c) => "%" + c.charCodeAt(0).toString(16))}`,
    );
    for (const name of ["content-length", "content-range", "accept-ranges"]) {
      const value = upstream.headers.get(name);
      if (value) headers.set(name, value);
    }
    return new Response(upstream.body, { status: upstream.status, headers });
  } catch (error) {
    const safe = publicError(error);
    return new Response(safe.message, { status: safe.status, headers });
  }
}

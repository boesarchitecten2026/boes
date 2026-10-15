import { readFile } from "node:fs/promises";
import path from "node:path";
import { PortalError, type Item } from "./graph";

// Deliberately separate from OneDrive: this project contains only fictional examples.
export const DEMO_PROJECT_ID = "zhq17sD-hJ324XCT1yKvb2IEsErxs_4c";
const project: Item = {
  id: "demo-root",
  name: "Voorbeeldproject — Woning aan de tuin",
  folder: { childCount: 3 },
};
const items: Item[] = [
  {
    id: "tekeningen",
    name: "Bouwtekeningen",
    folder: { childCount: 1 },
    parentReference: { id: project.id },
  },
  {
    id: "beelden",
    name: "Visualisaties",
    folder: { childCount: 1 },
    parentReference: { id: project.id },
  },
  {
    id: "toelichting",
    name: "Projecttoelichting.pdf",
    file: { mimeType: "application/pdf" },
    parentReference: { id: project.id },
  },
  {
    id: "plattegrond",
    name: "Begane grond.pdf",
    file: { mimeType: "application/pdf" },
    parentReference: { id: "tekeningen" },
  },
  {
    id: "impressie",
    name: "Schematische plattegrond.png",
    file: { mimeType: "image/png" },
    parentReference: { id: "beelden" },
  },
];
export function demoItem(id: string) {
  const item = [project, ...items].find((item) => item.id === id);
  if (!item) throw new PortalError(404, "Voorbeeldbestand niet gevonden.");
  return item;
}
export function demoProject(query: { folder?: string; preview?: string }) {
  const folder = query.folder ? demoItem(query.folder) : project;
  if (!folder.folder) throw new PortalError(404, "Voorbeeldmap niet gevonden.");
  return {
    project,
    folder,
    files: items.filter((item) => item.parentReference?.id === folder.id),
    preview: query.preview ? demoItem(query.preview) : undefined,
    isDemo: true,
  };
}
export async function demoContent(id: string) {
  // Explicit mapping; user-supplied paths are never read from disk.
  const files: Record<string, string> = {
    toelichting: "toelichting.pdf",
    plattegrond: "plattegrond.pdf",
    impressie: "plattegrond.png",
  };
  const filename = files[id];
  if (!filename) throw new PortalError(404, "Voorbeeldbestand niet gevonden.");
  const data = await readFile(
    path.join(process.cwd(), "public", "demo", filename),
  );
  return new Response(new Uint8Array(data), {
    headers: { "content-length": String(data.length) },
  });
}

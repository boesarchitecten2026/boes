/* eslint-disable @next/next/no-img-element -- Same-origin authorized streaming endpoint; no image optimizer cache. */
import type { Metadata } from "next";
import { notFound } from "next/navigation";
import {
  authorizedItem,
  children,
  PortalError,
  previewType,
  publicError,
  resolveProject,
} from "@/lib/graph";

import { DEMO_PROJECT_ID, demoProject } from "@/lib/demo";

export const dynamic = "force-dynamic";
export async function generateMetadata({
  params,
}: {
  params: Promise<{ projectId: string }>;
}): Promise<Metadata> {
  const { projectId } = await params;
  return { manifest: `/${encodeURIComponent(projectId)}/manifest.webmanifest` };
}
async function loadProject({
  params,
  searchParams,
}: {
  params: Promise<{ projectId: string }>;
  searchParams: Promise<{ folder?: string; preview?: string }>;
}) {
  const { projectId } = await params;
  const query = await searchParams;
  try {
    if (projectId === DEMO_PROJECT_ID) return demoProject(query);
    const project = await resolveProject(projectId);
    const folder = query.folder
      ? await authorizedItem(project, query.folder)
      : project;
    if (!folder.folder) throw new PortalError(404, "Map niet gevonden.");
    const files = (await children(folder.id)).sort(
      (a, b) =>
        Number(!!b.folder) - Number(!!a.folder) ||
        a.name.localeCompare(b.name, "nl"),
    );
    const preview = query.preview
      ? await authorizedItem(project, query.preview)
      : undefined;
    return { project, folder, files, preview, isDemo: false };
  } catch (error) {
    return { error: publicError(error) };
  }
}
export default async function ProjectPage(props: {
  params: Promise<{ projectId: string }>;
  searchParams: Promise<{ folder?: string; preview?: string }>;
}) {
  const data = await loadProject(props);
  if ("error" in data) {
    if (data.error!.status === 404) notFound();
    return (
      <main className="mx-auto max-w-2xl p-8">
        <h1 className="mb-4 text-xl font-semibold">Portaal niet beschikbaar</h1>
        <p>{data.error!.message}</p>
      </main>
    );
  }
  const { project, folder, files, preview } = data;
  const { projectId } = await props.params;
  const mime = preview && previewType(preview);
  const base = `/${projectId}`;
  const folderQuery =
    folder.id === project.id ? "" : `folder=${encodeURIComponent(folder.id)}&`;
  return (
    <main className="portal">
      <a className="skip-link" href="#documents">Naar de documenten</a>
      <header className="portal-header">
        <a className="brand" href={base} aria-label="Boes Architecten — projectmap">
          <span>Boes</span><span>Architecten</span>
        </a>
        <span className="portal-label">Klantenportaal</span>
      </header>
      {data.isDemo && (
        <p className="demo-notice">
          Demonstratie met fictieve documenten. Dit project is niet gekoppeld
          aan OneDrive.
        </p>
      )}
      <div className="project-heading">
        <p className="eyebrow">01 / Projectdocumenten</p>
        <h1>{project.name}</h1>
        <p className="intro">Alle documenten voor uw project, op één plek.</p>
      </div>
      {folder.id !== project.id && (
        <nav className="folder-navigation" aria-label="Mapnavigatie">
          <a href={base}>Projectmap</a>
          <a
            href={`${base}?folder=${encodeURIComponent(folder.parentReference!.id!)}`}
          >
            ← Bovenliggende map
          </a>
          <span aria-current="page">{folder.name}</span>
        </nav>
      )}
      <div className="document-heading" id="documents" tabIndex={-1}>
        <h2>{folder.id === project.id ? "Uw documenten" : folder.name}</h2>
        <span>{files.length} {files.length === 1 ? "onderdeel" : "onderdelen"}</span>
      </div>
      <ul className="document-list">
        {files.map((item) => (
          <li
            key={item.id}
            className="document-row"
          >
            <div className="document-name">
              {item.folder ? (
                <a href={`${base}?folder=${encodeURIComponent(item.id)}`}>
                  <span className="file-kind">Map</span>{item.name}<span className="folder-arrow" aria-hidden="true">↗</span>
                </a>
              ) : previewType(item) ? (
                <a
                  href={`${base}?${folderQuery}preview=${encodeURIComponent(item.id)}`}
                >
                  <span className="file-kind">{previewType(item) === "application/pdf" ? "PDF" : "Afbeelding"}</span>
                  {item.name}
                </a>
              ) : (
                <span><span className="file-kind">Bestand</span>{item.name}</span>
              )}
            </div>
            {item.file && (
              <a
                className="download-link"
                href={`${base}/file/${encodeURIComponent(item.id)}`}
              >
                Download<span className="sr-only"> {item.name}</span>
              </a>
            )}
          </li>
        ))}
        {files.length === 0 && (
          <li className="p-4 text-gray-600">Deze map is leeg.</li>
        )}
      </ul>
      {preview && (
        <section className="preview-panel" aria-label="Bestandsvoorbeeld">
          <div className="preview-heading">
            <h2 className="break-words font-semibold">{preview.name}</h2>
            <a href={`${base}?${folderQuery}`}>Sluiten</a>
          </div>
          {mime === "application/pdf" ? (
            <iframe
              title={preview.name}
              className="pdf-preview"
              src={`${base}/file/${encodeURIComponent(preview.id)}?preview=1`}
            />
          ) : mime ? (
            <img
              alt={preview.name}
              className="image-preview"
              src={`${base}/file/${encodeURIComponent(preview.id)}?preview=1`}
            />
          ) : (
            <p>Voor dit bestand is geen voorbeeld beschikbaar.</p>
          )}
          <a
            className="preview-download"
            href={`${base}/file/${encodeURIComponent(preview.id)}`}
          >
            Bestand downloaden
          </a>
        </section>
      )}
      <footer className="portal-footer"><span>Boes Architecten</span><span>Design for change.</span></footer>
    </main>
  );
}

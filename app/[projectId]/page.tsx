/* eslint-disable @next/next/no-img-element -- Same-origin authorized streaming endpoint; no image optimizer cache. */
import { notFound } from "next/navigation";
import {
  authorizedItem,
  children,
  PortalError,
  previewType,
  publicError,
  resolveProject,
} from "@/lib/graph";

export const dynamic = "force-dynamic";
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
    return { project, folder, files, preview };
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
    <main className="mx-auto max-w-4xl p-5 sm:p-10">
      <h1 className="mb-6 text-2xl font-semibold">{project.name}</h1>
      {folder.id !== project.id && (
        <nav className="mb-5 flex flex-wrap gap-5">
          <a href={base}>Projectmap</a>
          <a
            href={`${base}?folder=${encodeURIComponent(folder.parentReference!.id!)}`}
          >
            ← Bovenliggende map
          </a>
          <span>{folder.name}</span>
        </nav>
      )}
      <ul className="divide-y divide-gray-200 rounded border border-gray-200">
        {files.map((item) => (
          <li
            key={item.id}
            className="flex items-center justify-between gap-4 p-4"
          >
            <div className="min-w-0 break-words">
              {item.folder ? (
                <a href={`${base}?folder=${encodeURIComponent(item.id)}`}>
                  📁 {item.name}
                </a>
              ) : previewType(item) ? (
                <a
                  href={`${base}?${folderQuery}preview=${encodeURIComponent(item.id)}`}
                >
                  {item.name}
                </a>
              ) : (
                <span>{item.name}</span>
              )}
            </div>
            {item.file && (
              <a
                className="shrink-0 text-sm"
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
        <section className="mt-8" aria-label="Bestandsvoorbeeld">
          <div className="mb-3 flex justify-between gap-4">
            <h2 className="break-words font-semibold">{preview.name}</h2>
            <a href={`${base}?${folderQuery}`}>Sluiten</a>
          </div>
          {mime === "application/pdf" ? (
            <iframe
              title={preview.name}
              className="h-[70vh] w-full border"
              src={`${base}/file/${encodeURIComponent(preview.id)}?preview=1`}
            />
          ) : mime ? (
            <img
              alt={preview.name}
              className="max-h-[75vh] max-w-full object-contain"
              src={`${base}/file/${encodeURIComponent(preview.id)}?preview=1`}
            />
          ) : (
            <p>Voor dit bestand is geen voorbeeld beschikbaar.</p>
          )}
          <a
            className="mt-3 inline-block"
            href={`${base}/file/${encodeURIComponent(preview.id)}`}
          >
            Bestand downloaden
          </a>
        </section>
      )}
    </main>
  );
}

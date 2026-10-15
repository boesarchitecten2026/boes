import manifest from "@/lib/manifest";
import { DEMO_PROJECT_ID } from "@/lib/demo";
import { publicError, resolveProject } from "@/lib/graph";

export async function GET(
  _request: Request,
  context: { params: Promise<{ projectId: string }> },
) {
  const headers = {
    "Content-Type": "application/manifest+json",
    "Cache-Control": "private, no-store",
  };
  try {
    const { projectId } = await context.params;
    if (projectId !== DEMO_PROJECT_ID) await resolveProject(projectId);
    const projectUrl = `/${encodeURIComponent(projectId)}`;
    return Response.json(
      {
        ...manifest(),
        // A stable app identity and launch URL for each project, independent of folder navigation.
        id: projectUrl,
        start_url: projectUrl,
        scope: "/",
      },
      { headers },
    );
  } catch (error) {
    const safe = publicError(error);
    return new Response(safe.message, {
      status: safe.status,
      headers: { "Cache-Control": "private, no-store" },
    });
  }
}

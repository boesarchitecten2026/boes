import { beforeEach, test } from "node:test";
import assert from "node:assert/strict";
import {
  authorizedItem,
  children,
  parseFolderUrl,
  projectId,
  projects,
  resolveProject,
} from "../lib/graph.ts";

const hostname = "example-my.sharepoint.com";
const sitePath = "/personal/architect_example_nl";
const libraryPath = sitePath + "/Documents";
const folderPath = libraryPath + "/App";
const browserUrl = `https://${hostname}/my?id=${encodeURIComponent(folderPath)}&viewid=example`;
let requests: URL[] = [];
let sequence = 0;
beforeEach(() => {
  Object.assign(process.env, {
    AZURE_TENANT_ID: "tenant",
    AZURE_CLIENT_ID: "client",
    AZURE_CLIENT_SECRET: "mock",
    ONEDRIVE_DRIVE_ID: "",
    ONEDRIVE_ROOT_FOLDER_ID: "",
    ONEDRIVE_FOLDER_URL: browserUrl + `&test=${sequence++}`,
    PROJECT_LINK_SECRET: "a".repeat(64),
  });
  requests = [];
  globalThis.fetch = async (input, init) => {
    const url = new URL(String(input));
    requests.push(url);
    if (url.hostname === "login.microsoftonline.com")
      return Response.json({ access_token: "mock", expires_in: 3600 });
    assert.equal(url.hostname, "graph.microsoft.com");
    assert.equal(
      new Headers(init?.headers).get("authorization"),
      "Bearer mock",
    );
    const path = decodeURIComponent(url.pathname);
    if (path === `/v1.0/sites/${hostname}:${sitePath}`)
      return Response.json({ id: "site-id" });
    if (path === "/v1.0/sites/site-id/drive")
      return Response.json({
        id: "drive",
        webUrl: `https://${hostname}${libraryPath}`,
      });
    if (path === "/v1.0/drives/drive/root:/App")
      return Response.json({ id: "root", folder: { childCount: 1 } });
    if (path === "/v1.0/drives/drive/items/root/children")
      return Response.json({
        value: [{ id: "project", name: "Example", folder: { childCount: 1 } }],
      });
    if (path === "/v1.0/drives/drive/items/project")
      return Response.json({
        id: "project",
        folder: { childCount: 1 },
        parentReference: { id: "root", driveId: "drive" },
      });
    if (path === "/v1.0/drives/drive/items/file")
      return Response.json({
        id: "file",
        file: { mimeType: "application/pdf" },
        parentReference: { id: "project", driveId: "drive" },
      });
    if (path === "/v1.0/drives/drive/items/foreign")
      return Response.json({
        id: "foreign",
        file: { mimeType: "application/pdf" },
        parentReference: { id: "root", driveId: "drive" },
      });
    if (path === "/v1.0/drives/drive/items/root")
      return Response.json({ id: "root", folder: {} });
    throw new Error(`Unexpected Graph path: ${path}`);
  };
});
test("parses the OneDrive browser URL and direct SharePoint folder URLs", () => {
  assert.deepEqual(parseFolderUrl(browserUrl), {
    hostname,
    sitePath,
    folderPath,
  });
  assert.equal(
    parseFolderUrl(`https://${hostname}${folderPath}`).folderPath,
    folderPath,
  );
  assert.equal(
    parseFolderUrl(
      "https://example.sharepoint.com/sites/Studio/Shared%20Documents/App",
    ).sitePath,
    "/sites/Studio",
  );
  assert.equal(
    parseFolderUrl(
      "https://example.sharepoint.com/teams/Studio/Shared%20Documents/Forms/AllItems.aspx?id=%2Fteams%2FStudio%2FShared%20Documents%2FApp",
    ).folderPath,
    "/teams/Studio/Shared Documents/App",
  );
});
test("rejects non-Microsoft hosts, personal OneDrive, credentials and traversal", () => {
  for (const url of [
    "https://evil.example/my?id=" + encodeURIComponent(folderPath),
    "https://example.sharepoint.com.evil.test/sites/A/Documents/App",
    "http://example.sharepoint.com/sites/A/Documents/App",
    "https://onedrive.live.com/",
    "https://user:pass@example.sharepoint.com/sites/A/Documents/App",
    `https://${hostname}/my?id=${encodeURIComponent(libraryPath + "/../private")}`,
    "https://example.sharepoint.com/:f:/s/shared-token",
  ])
    assert.throws(() => parseFolderUrl(url), { status: 503 });
});
test("resolves IDs server-side and preserves project links and project boundaries", async () => {
  const list = await projects();
  assert.equal(list[0].projectId, projectId("project", "drive"));
  assert.equal((await resolveProject(list[0].projectId)).id, "project");
  assert.equal((await authorizedItem(list[0], "file")).id, "file");
  await assert.rejects(authorizedItem(list[0], "foreign"), { status: 404 });
  assert.equal(
    requests.filter((url) => url.pathname.startsWith(`/v1.0/sites/${hostname}`))
      .length,
    1,
  );
  assert.ok(
    requests.every((url) =>
      ["graph.microsoft.com", "login.microsoftonline.com"].includes(
        url.hostname,
      ),
    ),
  );
});
test("explicit IDs take precedence; partial IDs fail before any lookup", async () => {
  process.env.ONEDRIVE_DRIVE_ID = "drive";
  process.env.ONEDRIVE_ROOT_FOLDER_ID = "root";
  assert.equal((await projects())[0].id, "project");
  assert.ok(!requests.some((url) => url.pathname.startsWith("/v1.0/sites/")));
  process.env.ONEDRIVE_ROOT_FOLDER_ID = "";
  await assert.rejects(projects(), { status: 503 });
});
test("discovery denies missing site access without broadening permissions or exposing Microsoft errors", async () => {
  for (const status of [403, 404, 400]) {
    globalThis.fetch = async () =>
      Response.json(
        { error: { message: "sensitive Microsoft detail" } },
        { status },
      );
    await assert.rejects(projects(), (error) => {
      assert.equal((error as { status: number }).status, 503);
      assert.match((error as Error).message, /tenant.*leesrechten/);
      assert.ok(!(error as Error).message.includes("sensitive"));
      return true;
    });
  }
});
test("library mismatch and missing folders do not fall back to the entire drive", async () => {
  const original = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    if (String(input).includes("/sites/site-id/drive"))
      return Response.json({
        id: "drive",
        webUrl: `https://${hostname}${sitePath}/Other`,
      });
    return original(input, init);
  };
  await assert.rejects(children("root"), { status: 503 });
  globalThis.fetch = async (input, init) => {
    if (String(input).includes("/root:/App"))
      return new Response(null, { status: 404 });
    return original(input, init);
  };
  await assert.rejects(projects(), { status: 503 });
});

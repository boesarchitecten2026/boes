import { test, beforeEach } from "node:test";
import assert from "node:assert/strict";
import {
  authorizedItem,
  children,
  fileContent,
  getItem,
  previewType,
  projectId,
  projects,
  resolveProject,
  type Item,
} from "../lib/graph.ts";

Object.assign(process.env, {
  AZURE_TENANT_ID: "tenant",
  AZURE_CLIENT_ID: "client",
  AZURE_CLIENT_SECRET: "not-a-real-secret",
  ONEDRIVE_DRIVE_ID: "drive",
  ONEDRIVE_ROOT_FOLDER_ID: "root",
  PROJECT_LINK_SECRET: "a".repeat(64),
});
const folder = (id: string, parent: string): Item => ({
  id,
  name: id,
  folder: { childCount: 0 },
  parentReference: { id: parent, driveId: "drive" },
});
const items: Record<string, Item> = {
  root: folder("root", "outside"),
  a: folder("a", "root"),
  b: folder("b", "root"),
  sub: folder("sub", "a"),
  pdf: {
    id: "pdf",
    name: "plan.pdf",
    file: { mimeType: "application/pdf" },
    parentReference: { id: "sub", driveId: "drive" },
  },
  foreign: {
    id: "foreign",
    name: "private.pdf",
    file: { mimeType: "application/pdf" },
    parentReference: { id: "b", driveId: "drive" },
  },
  shortcut: { ...folder("shortcut", "a"), remoteItem: {} },
};
let requests: string[] = [];
beforeEach(() => {
  requests = [];
  globalThis.fetch = async (input, init) => {
    const url = String(input);
    requests.push(url);
    if (url.includes("login.microsoftonline.com"))
      return Response.json({ access_token: "mock-token", expires_in: 3600 });
    assert.equal(
      new Headers(init?.headers).get("Authorization"),
      "Bearer mock-token",
    );
    const parsed = new URL(url);
    const id = parsed.pathname.split("/")[5];
    if (parsed.pathname.endsWith("/content"))
      return new Response(null, {
        status: 302,
        headers: { location: "https://files.example.test/content" },
      });
    if (parsed.pathname.endsWith("/children")) {
      if (parsed.searchParams.has("page"))
        return Response.json({ value: [items.b] });
      return Response.json({
        value: [items.a, items.shortcut],
        "@odata.nextLink":
          "https://graph.microsoft.com/v1.0/drives/drive/items/root/children?page=2",
      });
    }
    return items[id]
      ? Response.json(items[id])
      : new Response(null, { status: 404 });
  };
});
test("stable cryptographic IDs, renamed folders do not change links", () => {
  const id = projectId("a");
  assert.match(id, /^[\w-]{43}$/);
  assert.equal(projectId("a"), id);
  assert.notEqual(projectId("b"), id);
  process.env.PROJECT_LINK_SECRET = "b".repeat(64);
  assert.notEqual(projectId("a"), id);
  process.env.PROJECT_LINK_SECRET = "a".repeat(64);
});
test("pagination, direct subfolders and project resolution", async () => {
  assert.deepEqual(
    (await projects()).map((p) => p.id),
    ["a", "b"],
  );
  assert.equal((await resolveProject(projectId("a"))).id, "a");
  await assert.rejects(resolveProject("bad"), { status: 404 });
  await assert.rejects(resolveProject(projectId("unknown")), { status: 404 });
});
test("nested file allowed; sibling project, root, shortcut, missing file denied", async () => {
  assert.equal((await authorizedItem(items.a, "pdf")).id, "pdf");
  for (const id of ["foreign", "root", "shortcut", "missing"])
    await assert.rejects(authorizedItem(items.a, id), { status: 404 });
});
test("ancestry cycles fail closed", async () => {
  items.cycle = folder("cycle", "cycle");
  await assert.rejects(authorizedItem(items.a, "cycle"), { status: 404 });
  delete items.cycle;
});
test("active formats cannot be previewed inline", () => {
  assert.equal(previewType(items.pdf), "application/pdf");
  for (const mimeType of [
    "text/html",
    "image/svg+xml",
    "application/javascript",
  ])
    assert.equal(
      previewType({ id: "x", name: "x", file: { mimeType } }),
      undefined,
    );
});
test("server follows content URL without leaking bearer token; forwards range", async () => {
  const graphFetch = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    if (String(input).startsWith("https://files.example.test")) {
      assert.equal(new Headers(init?.headers).get("Authorization"), null);
      assert.equal(new Headers(init?.headers).get("Range"), "bytes=0-3");
      return new Response("test", { status: 206 });
    }
    return graphFetch(input, init);
  };
  assert.equal(
    await (await fileContent(items.pdf, "bytes=0-3")).text(),
    "test",
  );
});
test("Graph failures and untrusted pagination destinations fail closed", async () => {
  const original = globalThis.fetch;
  globalThis.fetch = async (input, init) => {
    if (String(input).includes("/children"))
      return Response.json({
        value: [],
        "@odata.nextLink": "https://evil.example/v1.0/stolen",
      });
    return original(input, init);
  };
  await assert.rejects(children("root"), { status: 502 });
  assert.ok(!requests.some((url) => url.includes("evil.example")));
  globalThis.fetch = async () => new Response(null, { status: 403 });
  await assert.rejects(getItem("pdf"), { status: 503 });
});

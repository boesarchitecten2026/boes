import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createHmac } from "node:crypto";

test("production HTTP: project listing, navigation, preview, download and project isolation", async () => {
  const server = spawn(
    process.execPath,
    [
      "--import",
      "./tests/fixtures/graph.mjs",
      "node_modules/next/dist/bin/next",
      "start",
      "-p",
      "3219",
    ],
    {
      env: {
        ...process.env,
        AZURE_TENANT_ID: "tenant",
        AZURE_CLIENT_ID: "client",
        AZURE_CLIENT_SECRET: "mock",
        ONEDRIVE_DRIVE_ID: "drive",
        ONEDRIVE_ROOT_FOLDER_ID: "root",
        PROJECT_LINK_SECRET: "a".repeat(64),
      },
      stdio: ["ignore", "pipe", "pipe"],
    },
  );
  let log = "";
  server.stdout.on("data", (chunk) => {
    log += chunk;
  });
  server.stderr.on("data", (chunk) => {
    log += chunk;
  });
  const base = "http://localhost:3219";
  try {
    let ready = false;
    for (let i = 0; i < 100; i++) {
      try {
        await fetch(base);
        ready = true;
        break;
      } catch {
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
    }
    assert.ok(ready, log);
    const id = createHmac("sha256", "a".repeat(64))
      .update("drive:a")
      .digest("base64url");
    const root = await fetch(base);
    assert.equal(root.status, 200);
    assert.ok(!(await root.text()).includes("Project voorbeeld"));
    const project = await fetch(`${base}/${id}`);
    assert.equal(project.status, 200);
    assert.match(await project.text(), /Project voorbeeld/);
    const sub = await fetch(`${base}/${id}?folder=sub`);
    assert.match(await sub.text(), /Tekening.pdf/);
    const preview = await fetch(`${base}/${id}?preview=pdf`);
    assert.match(await preview.text(), /iframe/);
    const file = await fetch(`${base}/${id}/file/pdf?preview=1`);
    assert.equal(file.status, 200);
    assert.equal(file.headers.get("content-type"), "application/pdf");
    assert.equal(await file.text(), "%PDF-mock");
    assert.match(file.headers.get("cache-control")!, /no-store/);
    assert.equal(file.headers.get("location"), null);
    const download = await fetch(`${base}/${id}/file/pdf`);
    assert.match(download.headers.get("content-disposition")!, /^attachment/);
    for (const path of [
      `/${id}/file/foreign`,
      `/${id}?folder=b`,
      `/${id}?preview=foreign`,
      "/invalid",
    ])
      assert.equal((await fetch(base + path)).status, 404, path);
    assert.equal(
      (await fetch(`${base}/${id}/file/pdf`, { method: "POST" })).status,
      405,
    );
    assert.equal((await fetch(base + "/manifest.webmanifest")).status, 200);
    assert.equal((await fetch(base + "/icons/icon-192.png")).status, 200);
  } finally {
    server.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      if (server.exitCode !== null) resolve();
      else server.once("exit", () => resolve());
    });
  }
});

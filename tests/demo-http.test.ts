import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";

test("demo works without Microsoft credentials, with folders, previews and real example files", async () => {
  const server = spawn(
    process.execPath,
    ["node_modules/next/dist/bin/next", "start", "-p", "3220"],
    {
      env: {
        ...process.env,
        AZURE_TENANT_ID: "",
        AZURE_CLIENT_ID: "",
        AZURE_CLIENT_SECRET: "",
        ONEDRIVE_DRIVE_ID: "",
        ONEDRIVE_ROOT_FOLDER_ID: "",
        PROJECT_LINK_SECRET: "",
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
  const base = "http://localhost:3220/zhq17sD-hJ324XCT1yKvb2IEsErxs_4c";
  try {
    let ready = false;
    for (let i = 0; i < 100; i++) {
      try {
        await fetch("http://localhost:3220");
        ready = true;
        break;
      } catch {
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
    }
    assert.ok(ready, log);
    const page = await fetch(base);
    assert.equal(page.status, 200);
    assert.match(await page.text(), /Demonstratie met fictieve documenten/);
    const folder = await fetch(base + "?folder=tekeningen");
    assert.match(await folder.text(), /Begane grond.pdf/);
    const pdfPage = await fetch(
      base + "?folder=tekeningen&preview=plattegrond",
    );
    assert.match(await pdfPage.text(), /iframe/);
    const imagePage = await fetch(base + "?folder=beelden&preview=impressie");
    assert.match(await imagePage.text(), /<img/);
    const pdf = await fetch(base + "/file/plattegrond?preview=1");
    assert.equal(pdf.headers.get("content-type"), "application/pdf");
    assert.match(await pdf.text(), /^%PDF-1.4/);
    const png = await fetch(base + "/file/impressie?preview=1");
    assert.equal(png.headers.get("content-type"), "image/png");
    assert.deepEqual(
      [...new Uint8Array(await png.arrayBuffer()).slice(0, 8)],
      [137, 80, 78, 71, 13, 10, 26, 10],
    );
    const download = await fetch(base + "/file/toelichting");
    assert.match(download.headers.get("content-disposition")!, /^attachment/);
    for (const path of [
      "?folder=unknown",
      "/file/unknown",
      "?preview=unknown",
      "/file/tekeningen",
    ])
      assert.equal((await fetch(base + path)).status, 404);
    const real = await fetch("http://localhost:3220/" + "a".repeat(43));
    assert.match(await real.text(), /niet geconfigureerd/);
  } finally {
    server.kill("SIGTERM");
    await new Promise<void>((resolve) => {
      if (server.exitCode !== null) resolve();
      else server.once("exit", () => resolve());
    });
  }
});

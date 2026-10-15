// Only loaded by the HTTP test subprocess. No production mock mode exists.
const original = globalThis.fetch;
const folder = (id, parent) => ({
  id,
  name: id === "a" ? "Project voorbeeld" : id,
  folder: { childCount: 1 },
  parentReference: { id: parent, driveId: "drive" },
});
const items = {
  a: folder("a", "root"),
  b: folder("b", "root"),
  sub: folder("sub", "a"),
  pdf: {
    id: "pdf",
    name: "Tekening.pdf",
    file: { mimeType: "application/pdf" },
    parentReference: { id: "sub", driveId: "drive" },
  },
  foreign: {
    id: "foreign",
    name: "Privé.pdf",
    file: { mimeType: "application/pdf" },
    parentReference: { id: "b", driveId: "drive" },
  },
};
globalThis.fetch = async (input, init) => {
  const url = new URL(
    typeof input === "string" || input instanceof URL ? input : input.url,
  );
  if (url.hostname === "login.microsoftonline.com")
    return Response.json({ access_token: "mock", expires_in: 3600 });
  if (url.hostname === "files.example.test")
    return new Response("%PDF-mock", { headers: { "content-length": "9" } });
  if (url.hostname !== "graph.microsoft.com") return original(input, init);
  const id = url.pathname.split("/")[5];
  if (url.pathname.endsWith("/content"))
    return new Response(null, {
      status: 302,
      headers: { location: "https://files.example.test/content" },
    });
  if (url.pathname.endsWith("/children"))
    return Response.json({
      value: Object.values(items).filter(
        (item) => item.parentReference.id === id,
      ),
    });
  return items[id]
    ? Response.json(items[id])
    : new Response(null, { status: 404 });
};

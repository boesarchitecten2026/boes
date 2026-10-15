import { projects, publicError } from "../lib/graph.ts";
try {
  const base = new URL(process.env.APP_BASE_URL || "http://localhost:3000");
  if (!["https:", "http:"].includes(base.protocol))
    throw new Error("Ongeldige APP_BASE_URL");
  for (const project of await projects())
    console.log(
      `${project.name}\t${new URL("/" + project.projectId, base).href}`,
    );
} catch (error) {
  console.error(publicError(error).message);
  process.exitCode = 1;
}

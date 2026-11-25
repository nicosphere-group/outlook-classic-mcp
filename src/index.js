// Entry point: an MCP server over stdio. stdout carries the protocol; logs go to stderr.

import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

import { AuditLog } from "./audit.js";
import { loadConfig } from "./config.js";
import { DialogConfirmer, PowerShellBackend } from "./powershellBackend.js";
import { buildServer } from "./server.js";
import { MailService } from "./service.js";

async function main() {
  if (process.platform !== "win32") {
    console.error("outlook-classic-mcp works only on Windows with classic Outlook installed.");
    process.exit(1);
  }
  const config = loadConfig();
  const backend = new PowerShellBackend();
  const service = new MailService(backend, new DialogConfirmer(backend, config.lang), new AuditLog(config.auditDir), config);
  const server = buildServer(service, config);

  const shutdown = () => {
    backend.close();
    process.exit(0);
  };
  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
  process.stdin.on("close", shutdown);

  await server.connect(new StdioServerTransport());
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

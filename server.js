"use strict";

const fs = require("node:fs");
const http = require("node:http");
const path = require("node:path");

const HOST = "127.0.0.1";
const PORT = Number(process.env.PORT || 8001);
const DEFAULT_FOLDER = process.env.FRESHDESK_FOLDER || String.raw`C:\Users\soumya.roy\OneDrive - SHYAM STEEL INDUSTRIES LIMITED\Projects\#Freshdesk`;
const ROOT = __dirname;
const DUMP_FILENAME = /^FreshDesk_CRM_.*\.xlsx?$/i;
const STATIC_FILES = new Set(["index.html", "styles.css", "app.js"]);

function sendJson(response, status, body) {
  response.writeHead(status, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
  response.end(JSON.stringify(body));
}

async function handleRequest(request, response) {
  if (request.method !== "GET") {
    response.writeHead(405, { Allow: "GET" });
    response.end("Method not allowed");
    return;
  }

  const url = new URL(request.url, `http://${HOST}:${PORT}`);
  if (url.pathname === "/api/default-folder") {
    try {
      const entries = await fs.promises.readdir(DEFAULT_FOLDER, { withFileTypes: true });
      const files = entries
        .filter(entry => entry.isFile() && DUMP_FILENAME.test(entry.name))
        .map(entry => entry.name)
        .sort((left, right) => left.localeCompare(right));
      sendJson(response, 200, { folder: DEFAULT_FOLDER, files });
    } catch (error) {
      sendJson(response, 404, { error: `Cannot read the configured Freshdesk folder: ${error.message}`, folder: DEFAULT_FOLDER });
    }
    return;
  }

  if (url.pathname.startsWith("/api/files/")) {
    let fileName;
    try {
      fileName = decodeURIComponent(url.pathname.slice("/api/files/".length));
    } catch {
      sendJson(response, 400, { error: "Invalid file name." });
      return;
    }
    if (!fileName || path.basename(fileName) !== fileName || !DUMP_FILENAME.test(fileName)) {
      sendJson(response, 400, { error: "Invalid Freshdesk export file name." });
      return;
    }
    try {
      const filePath = path.join(DEFAULT_FOLDER, fileName);
      const stats = await fs.promises.stat(filePath);
      if (!stats.isFile()) throw new Error("Not a file");
      response.writeHead(200, {
        "Content-Type": "application/octet-stream",
        "Content-Length": stats.size,
        "Cache-Control": "no-store"
      });
      fs.createReadStream(filePath).pipe(response);
    } catch {
      sendJson(response, 404, { error: "Freshdesk export not found." });
    }
    return;
  }

  const fileName = url.pathname === "/" ? "index.html" : url.pathname.slice(1);
  if (!STATIC_FILES.has(fileName)) {
    response.writeHead(404);
    response.end("Not found");
    return;
  }
  const contentTypes = { "index.html": "text/html; charset=utf-8", "styles.css": "text/css; charset=utf-8", "app.js": "text/javascript; charset=utf-8" };
  fs.createReadStream(path.join(ROOT, fileName))
    .on("error", () => {
      response.writeHead(500);
      response.end("Could not read application file.");
    })
    .on("open", () => response.writeHead(200, { "Content-Type": contentTypes[fileName] }))
    .pipe(response);
}

const server = http.createServer((request, response) => {
  handleRequest(request, response).catch(error => {
    console.error(error);
    if (!response.headersSent) response.writeHead(500);
    response.end("Request failed.");
  });
});

server.listen(PORT, HOST, () => {
  console.log(`Freshdesk Snapshot Analyzer: http://${HOST}:${PORT}`);
  console.log(`Default folder: ${DEFAULT_FOLDER}`);
});
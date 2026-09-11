"use strict";

const assert = require("assert");
const path = require("path");
const { spawn } = require("child_process");

const port = 3217;
const baseUrl = "http://127.0.0.1:" + port;
const child = spawn(process.execPath, ["server.js"], {
  cwd: __dirname,
  env: { ...process.env, PORT: String(port), OPENAI_API_KEY: "" },
  stdio: ["ignore", "pipe", "pipe"],
  windowsHide: true
});

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitForServer() {
  for (let attempt = 0; attempt < 30; attempt++) {
    try {
      const response = await fetch(baseUrl + "/api/health");
      if (response.ok) return response.json();
    } catch {
      // Server may still be starting.
    }
    await delay(100);
  }
  throw new Error("本機伺服器未在期限內啟動");
}

(async () => {
  try {
    const health = await waitForServer();
    assert.equal(health.ok, true);
    assert.equal(health.apiKeyConfigured, false);

    const page = await fetch(baseUrl + "/");
    const html = await page.text();
    assert.equal(page.status, 200);
    assert.match(html, /async function askAI/);

    const chat = await fetch(baseUrl + "/api/chat", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ message: "測試" })
    });
    const chatBody = await chat.json();
    assert.equal(chat.status, 503);
    assert.match(chatBody.error, /OPENAI_API_KEY/);

    console.log("Server smoke test: OK");
  } finally {
    child.kill();
  }
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

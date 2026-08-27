/**
 * İnci Akü PPWR — Electron shell.
 * Loads the Vite UI (5173). API must be on 8791 (started by CMD).
 */
const { app, BrowserWindow, shell } = require("electron");
const path = require("path");

const UI = process.env.INCI_PPWR_UI_URL || "http://127.0.0.1:5173";
const API_HEALTH = process.env.INCI_PPWR_API_HEALTH || "http://127.0.0.1:8791/api/health";
const WAIT_MS = Number(process.env.INCI_PPWR_WAIT_MS || 90000);

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function waitUrl(url, label) {
  const t0 = Date.now();
  while (Date.now() - t0 < WAIT_MS) {
    try {
      const res = await fetch(url, { method: "GET" });
      if (res.ok) return true;
    } catch {
      /* retry */
    }
    await sleep(500);
  }
  throw new Error(`Timeout waiting for ${label}: ${url}`);
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 860,
    minWidth: 960,
    minHeight: 640,
    title: "İnci Akü · PPWR Compliance Suite",
    icon: path.join(__dirname, "inci-aku-logo.png"),
    backgroundColor: "#070b14",
    autoHideMenuBar: true,
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: "deny" };
  });

  win.loadURL(UI);
  return win;
}

app.whenReady().then(async () => {
  try {
    await waitUrl(API_HEALTH, "API");
    await waitUrl(UI, "UI");
  } catch (err) {
    console.error(String(err));
    const fail = new BrowserWindow({
      width: 640,
      height: 320,
      title: "İnci Akü PPWR — Başlatılamadı",
      autoHideMenuBar: true,
    });
    fail.loadURL(
      "data:text/html;charset=utf-8," +
        encodeURIComponent(
          `<!doctype html><html><body style="font-family:Segoe UI,sans-serif;padding:2rem;background:#0E2A47;color:#fff">
          <h1>Başlatılamadı</h1>
          <p>${String(err).replace(/</g, "&lt;")}</p>
          <p>Önce <code>00_START_PPWR_DESKTOP.cmd</code> ile API + UI’yi açın.</p>
          </body></html>`,
        ),
    );
    return;
  }

  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

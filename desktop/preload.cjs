const { contextBridge } = require("electron");

contextBridge.exposeInMainWorld("inciPpwrDesktop", {
  isDesktop: true,
  shell: "electron",
});

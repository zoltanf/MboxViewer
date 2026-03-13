const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("mboxApi", {
  openMbox: () => ipcRenderer.invoke("open-mbox"),
  searchMessages: (payload) => ipcRenderer.invoke("search-messages", payload),
  getMessage: (payload) => ipcRenderer.invoke("get-message", payload),
  openExternal: (payload) => ipcRenderer.invoke("open-external", payload),
  copyToClipboard: (payload) => ipcRenderer.invoke("copy-to-clipboard", payload),
  openAttachmentPreview: (payload) => ipcRenderer.invoke("open-attachment-preview", payload),
  saveAttachment: (payload) => ipcRenderer.invoke("save-attachment", payload),
  saveMessageEml: (payload) => ipcRenderer.invoke("save-message-eml", payload),
  onIndexProgress: (callback) => {
    if (typeof callback !== "function") {
      return () => {};
    }

    const listener = (_, payload) => callback(payload);
    ipcRenderer.on("mbox-index-progress", listener);
    return () => {
      ipcRenderer.removeListener("mbox-index-progress", listener);
    };
  }
});

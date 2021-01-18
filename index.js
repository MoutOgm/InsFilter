const { app, BrowserWindow, Menu, globalShortcut } = require('electron')
function createWindow (x, y, t) {
  // Create the browser window.
  const win = new BrowserWindow({
    width: x,
    height: y,
    webPreferences: {
      nodeIntegration: true
    }
  })
  // and load the index.html of the app.
  win.menuBarVisible = false
  win.loadFile(t)
}

app.whenReady().then(load)

function load() {
  createWindow(800, 300, "Copy.html")
  createWindow(850, 625, "index.html")
}

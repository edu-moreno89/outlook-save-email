const { app, dialog, BrowserWindow } = require('electron');
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const path = require('path');
const fs = require('fs');

let mainWindow;
const serverApp = express();
const PORT = 3001;

serverApp.use(cors());
serverApp.use(bodyParser.json({ limit: '20mb' }));

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 400,
    height: 300,
    show: false,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  // Optional HTML content for the window
  mainWindow.loadURL('data:text/html,<h1>Folder Helper App Running</h1>');
}

app.whenReady().then(() => {
  createWindow();

  serverApp.get('/select-folder', async (req, res) => {
    console.log('select-folder');
    try {
      mainWindow.show();

      const result = await dialog.showOpenDialog(mainWindow, {
        title: "Select Folder to Save Attachments",
        properties: ['openDirectory']
      });

      // mainWindow.hide();
      console.log('Dialog result:', result);

      if (result.canceled || !result.filePaths.length) {        
        res.send('');
      } else {
        const path = result.filePaths[0];
        console.log('path:', path);
        res.send(path);
      }
    } catch (err) {
      console.error('Error showing dialog:', err);
      res.status(500).send('Failed to open dialog');
    }
  });

  serverApp.post('/save-file', (req, res) => {
    const { fileName, folderPath, base64Data } = req.body;
    console.log('save-file', fileName, base64Data.length);
  
    const filePath = path.join(folderPath, fileName);
    const buffer = Buffer.from(base64Data, 'base64');
    fs.writeFileSync(filePath, buffer);

    console.log(`File saved to ${filePath}`);
  });

  serverApp.listen(PORT, () => {
    console.log(`Folder Picker Server running at http://localhost:${PORT}`);
  });
});

app.on('window-all-closed', () => {
  app.quit();
});

const { app, BrowserWindow, globalShortcut } = require('electron');
process.env['ELECTRON_DISABLE_SECURITY_WARNINGS'] = 'true';

let mainWindow;

function createWindow() 
{
    mainWindow = new BrowserWindow({ width: 1300, height: 1000 });

    mainWindow.loadFile('index.html');

    mainWindow.webContents.openDevTools()
    mainWindow.on('closed', () => {
        mainWindow = null;
    });

    globalShortcut.register('CommandOrControl+R', () => {
        console.log('Reloading app!');
        mainWindow.reload();
    });    
}

app.on('ready', createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (mainWindow === null) {
        createWindow();
    }
});

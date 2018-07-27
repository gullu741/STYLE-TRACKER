const electron = require('electron')

const {app} = electron;//control app life

const {BrowserWindow} = electron;//create native browser window
const {Menu} = electron;
const {dialog} = electron; // dialog box creator
const {ipcMain} = electron;// inter process comunication reciever
let win

var dialogProperties = {
    filters : [
        {name:'Excel Files',extensions:['xlsm','xlsx']}
    ]
}

ipcMain.on('openFile',(event,path)=>{
    
    dialog.showOpenDialog(dialogProperties,function(filenames){
        event.sender.send("filename",filenames)
    })
})

function createWindow() {
    win = new BrowserWindow({width:1000,height:1000});

    win.loadURL(`file://${__dirname}/index.html`);
    
    // win.webContents.openDevTools();
    
    win.on('closed',()=>{
        win = null;
    })
    

}

const template = []

app.on('ready',()=>{
    createWindow();
    Menu.setApplicationMenu(Menu.buildFromTemplate(template))

});

app.on('window-all-closed',()=>{
    if(process.platform !== 'darwin'){
        app.quit();
    }
});


app.on('activate',()=>{
    if(win == null){
        createWindow();
    }
});
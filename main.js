const electron = require('electron')
const { BrowserWindow, Menu, dialog } = electron
const app = electron.app
const path = require('path')

let mainWindow

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    title: '劳动案件管理系统',
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      webSecurity: true
    }
  })

  mainWindow.loadFile('index.html')

  // Minimal app menu — no dev tools exposed in production
  const menu = Menu.buildFromTemplate([
    {
      label: '文件',
      submenu: [
        { label: '退出', accelerator: 'Alt+F4', click: () => app.quit() }
      ]
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '关于',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: '关于',
              message: '劳动案件管理系统',
              detail: `版本：${app.getVersion()}\n\n律师个人案件管理桌面应用\n数据保存位置：%APPDATA%\\劳动案件管理系统`
            })
          }
        }
      ]
    }
  ])
  Menu.setApplicationMenu(menu)
}

app.whenReady().then(() => {
  createWindow()

  // Only enable auto-updater in packaged app (not during development)
  if (app.isPackaged) {
    try {
      const { autoUpdater } = require('electron-updater')

      setTimeout(() => {
        autoUpdater.checkForUpdatesAndNotify().catch(err => {
          console.error('Update check failed:', err)
        })
      }, 3000)

      autoUpdater.on('update-available', () => {
        dialog.showMessageBox(mainWindow, {
          type: 'info',
          title: '发现新版本',
          message: '发现新版本，正在后台下载，下载完成后将提示重启。'
        })
      })

      autoUpdater.on('update-downloaded', () => {
        const response = dialog.showMessageBoxSync(mainWindow, {
          type: 'info',
          title: '更新已就绪',
          message: '新版本已下载完成，重启应用以完成更新。',
          buttons: ['立即重启', '稍后'],
          defaultId: 0
        })
        if (response === 0) {
          autoUpdater.quitAndInstall()
        }
      })

      autoUpdater.on('error', (err) => {
        // Silently log — don't interrupt the user for network errors
        console.error('Auto-updater error:', err)
      })
    } catch (err) {
      console.error('Failed to initialize auto-updater:', err)
    }
  }
})

app.on('window-all-closed', () => {
  app.quit()
})

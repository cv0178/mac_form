import { app, BrowserWindow, dialog, ipcMain, Menu, Notification, shell } from 'electron';
import path from 'node:path';
import { existsSync } from 'node:fs';
import { mkdir } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import started from 'electron-squirrel-startup';


if (started) app.quit();

const knowledgeFilters = [{ name: '资料文件（Word、Excel、PDF）', extensions: ['docx', 'xlsx', 'pdf'] }];
const formFilters = [{ name: '待填写表单（Word、Excel）', extensions: ['docx', 'xlsx'] }];

function toFileInfo(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  return {
    path: filePath,
    name: path.basename(filePath),
    extension,
    kind: { '.docx': 'Word 文档', '.xlsx': 'Excel 表格', '.pdf': 'PDF 文档' }[extension] ?? '文件',
  };
}

async function getAppDataPaths() {
  const root = path.join(app.getPath('userData'), 'data');
  const paths = {
    root,
    knowledge: path.join(root, 'knowledge'),
    forms: path.join(root, 'forms'),
    outputs: path.join(root, 'outputs'),
  };
  await Promise.all(Object.values(paths).map((directory) => mkdir(directory, { recursive: true })));
  return paths;
}

const BACKEND_URL = 'http://127.0.0.1:8999';
let backendProcess = null;

const wait = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));

async function waitForBackend() {
  for (let attempt = 0; attempt < 40; attempt += 1) {
    try {
      const response = await fetch(`${BACKEND_URL}/health`);
      if (response.ok) return;
    } catch {
      // 后端还在启动，继续等待。
    }
    await wait(250);
  }
  throw new Error('本地填写服务启动超时，请重新安装应用或联系技术支持。');
}

async function startPackagedBackend(appPaths) {
  // 开发时保持原有工作流：由 PyCharm 手动运行 main.py，避免抢占 8999 端口。
  if (!app.isPackaged) return;
  // Windows 上 PyInstaller 产物带 .exe 后缀，macOS / Linux 上没有后缀。
  const backendBinary = process.platform === 'win32' ? 'fill_form_backend.exe' : 'fill_form_backend';
  const backendPath = path.join(process.resourcesPath, 'backend', backendBinary);
  if (!existsSync(backendPath)) {
    throw new Error(`未找到内置后端程序：${backendPath}`);
  }
  backendProcess = spawn(backendPath, [], {
    cwd: path.dirname(backendPath),
    windowsHide: true,
    env: { ...process.env, FILL_FORM_DATA_DIR: appPaths.root },
    stdio: ['ignore', 'pipe', 'pipe'],
  });
  backendProcess.stderr.on('data', (data) => console.error(`[填表后端] ${data}`));
  backendProcess.on('error', (error) => console.error('无法启动内置填写服务：', error));
  await waitForBackend();
}

function stopPackagedBackend() {
  if (backendProcess && !backendProcess.killed) backendProcess.kill();
  backendProcess = null;
}

async function backendRequest(pathname, options) {
  let response;
  try {
    response = await fetch(`${BACKEND_URL}${pathname}`, options);
  } catch {
    throw new Error('无法连接本地服务，请确认 Python 后端正在运行');
  }
  if (!response.ok) throw new Error(`本地服务返回错误（${response.status}）`);
  return response.status === 204 ? null : response.json();
}

async function ensureOutputPath(filePath) {
  const resolved = path.resolve(filePath);
  // 导出目录可以由用户在客户端配置，所以校验基准要向后端要，不能写死
  const paths = await backendRequest('/api/system/paths');
  const relativePath = path.relative(paths.outputs, resolved);
  if (!relativePath.startsWith('..') && !path.isAbsolute(relativePath)) return resolved;

  // 改过导出目录之后，历史记录仍指向旧目录，所以再确认一次它是不是登记过的导出文件
  const result = await backendRequest('/api/exports/verify', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ path: resolved }),
  });
  if (!result.allowed) throw new Error('无权打开此文件位置');
  return resolved;
}

function registerIpcHandlers() {
  ipcMain.handle('desktop:get-app-data-paths', getAppDataPaths);
  ipcMain.handle('desktop:select-knowledge-files', async () => {
    const result = await dialog.showOpenDialog({ title: '选择个人资料', properties: ['openFile', 'multiSelections'], filters: knowledgeFilters });
    return result.canceled ? [] : result.filePaths.map(toFileInfo);
  });
  ipcMain.handle('desktop:select-directory', async () => {
    const result = await dialog.showOpenDialog({ title: '选择导出文件保存位置', properties: ['openDirectory', 'createDirectory'] });
    return result.canceled || result.filePaths.length === 0 ? null : result.filePaths[0];
  });
  ipcMain.handle('desktop:show-output-in-folder', async (_event, filePath) => {
    shell.showItemInFolder(await ensureOutputPath(filePath));
  });
  ipcMain.handle('desktop:open-output-file', async (_event, filePath) => {
    const errorMessage = await shell.openPath(await ensureOutputPath(filePath));
    if (errorMessage) throw new Error(errorMessage);
  });
  ipcMain.handle('desktop:notify-export-complete', (event, outputName) => {
    if (!Notification.isSupported()) return { sent: false, reason: '当前系统不支持桌面通知' };
    const notification = new Notification({
      title: '智填助手',
      body: `填写完成，已生成：${path.basename(String(outputName || '填写结果文件'))}`,
    });
    notification.on('click', () => {
      const window = BrowserWindow.fromWebContents(event.sender);
      if (window) {
        if (window.isMinimized()) window.restore();
        window.show();
        window.focus();
      }
    });
    notification.show();
    return { sent: true };
  });
  ipcMain.handle('desktop:select-form-file', async () => {
    const result = await dialog.showOpenDialog({ title: '选择待填写表单', properties: ['openFile'], filters: formFilters });
    return result.canceled || result.filePaths.length === 0 ? null : toFileInfo(result.filePaths[0]);
  });
}

const createWindow = () => {
  const mainWindow = new BrowserWindow({
    width: 1080, height: 720, minWidth: 800, minHeight: 560,
    webPreferences: { preload: path.join(__dirname, 'preload.js'), contextIsolation: true, nodeIntegration: false },
  });
  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    mainWindow.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL);
//    mainWindow.webContents.openDevTools();        // 启动时自动打开开发者工具
  } else {
    mainWindow.loadFile(path.join(__dirname, `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`));
  }
};

app.whenReady().then(async () => {
  Menu.setApplicationMenu(null);   // 去掉顶部菜单栏
  const appPaths = await getAppDataPaths();
  try {
    await startPackagedBackend(appPaths);
  } catch (error) {
    await dialog.showMessageBox({ type: 'error', title: '智填助手启动失败', message: error.message });
    app.quit();
    return;
  }
  registerIpcHandlers();
  createWindow();
  app.on('activate', () => { if (BrowserWindow.getAllWindows().length === 0) createWindow(); });
});
app.on('before-quit', stopPackagedBackend);
app.on('window-all-closed', () => { if (process.platform !== 'darwin') app.quit(); });


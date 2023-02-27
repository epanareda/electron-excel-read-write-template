// See the Electron documentation for details on how to use preload scripts:
// https://www.electronjs.org/docs/latest/tutorial/process-model#preload-scripts
const XLSX = require("xlsx");

const { contextBridge } = require('electron')

contextBridge.exposeInMainWorld('XLSX', 
  XLSX,
  // we can also expose variables, not just functions
)
// contextBridge.exposeInMainWorld('library name', 
//   Library,
// )
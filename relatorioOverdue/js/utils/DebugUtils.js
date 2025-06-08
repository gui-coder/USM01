// js/utils/DebugUtils.js
export class DebugUtils {
  static logError(error) {
    try {
      const errorLog = document.getElementById("errorLog");
      if (errorLog) {
        errorLog.innerHTML += `<p>${error}</p>`;
      }
      console.error(error);
    } catch (e) {
      console.error("Erro ao registrar erro:", e);
    }
  }

  static addDebugInfo(message) {
    try {
      const debugDiv = document.getElementById("debugInfo");
      if (debugDiv) {
        const timestamp = new Date().toLocaleTimeString();
        debugDiv.innerHTML += `<p>[${timestamp}] ${message}</p>`;
      }
    } catch (e) {
      console.error("Erro ao adicionar informação de debug:", e);
    }
  }

  static clearDebugInfo() {
    try {
      const debugDiv = document.getElementById("debugInfo");
      if (debugDiv) {
        debugDiv.innerHTML = "<h3>Informações de Debug</h3>";
      }
      const errorLog = document.getElementById("errorLog");
      if (errorLog) {
        errorLog.innerHTML = "";
      }
    } catch (e) {
      console.error("Erro ao limpar informações de debug:", e);
    }
  }
}

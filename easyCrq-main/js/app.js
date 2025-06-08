// js/app.js
import { readExcelFile } from "./fileReader.js";
import ExcelProcessor from "./excelProcessor.js";

const App = {
  // Estado inicial da aplicação
  state: {
    files: [],
    isProcessing: false,
    successCount: 0,
    errorCount: 0,
  },

  // Elementos do DOM
  elements: {},

  // Inicialização
  init() {
    this.initializeElements();
    this.bindEvents();
  },

  // Inicializa referências aos elementos do DOM
  initializeElements() {
    this.elements = {
      fileInput: document.getElementById("fileInput"),
      dropZone: document.getElementById("dropZone"),
      loadingIndicator: document.getElementById("loadingIndicator"),
      progressBar: document.querySelector(".progress-bar"),
      resultsContainer: document.getElementById("results"),
      fileList: document.getElementById("fileList"),
    };

    // Inicializa área de drop
    if (this.elements.dropZone) {
      this.initializeDropZone();
    }
  },

  // Vincula eventos aos elementos
  bindEvents() {
    this.elements.fileInput?.addEventListener("change", (e) =>
      this.handleFileSelect(e)
    );
  },

  // Inicializa área de drag and drop
  initializeDropZone() {
    const dropZone = this.elements.dropZone;

    dropZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      dropZone.classList.add("drag-over");
    });

    dropZone.addEventListener("dragleave", () => {
      dropZone.classList.remove("drag-over");
    });

    dropZone.addEventListener("drop", (e) => {
      e.preventDefault();
      dropZone.classList.remove("drag-over");
      this.handleFileSelect({ target: { files: e.dataTransfer.files } });
    });
  },

  // Manipulação de arquivos
  async handleFileSelect(event) {
    const files = event.target.files;
    if (!files?.length) return;

    this.state.files = Array.from(files);
    this.updateUI("processing");

    try {
      await this.processFiles();
      this.showNotification("success", "Arquivos processados com sucesso!");
    } catch (error) {
      this.showError("Erro ao processar arquivos: " + error.message);
    } finally {
      this.updateUI("completed");
    }
  },

  // Atualiza interface baseado no estado
  updateUI(status) {
    switch (status) {
      case "processing":
        this.clearResults();
        this.showLoading(true);
        this.updateFileList();
        break;
      case "completed":
        this.showLoading(false);
        this.updateProgress(100);
        break;
    }
  },

  // Processa os arquivos
  async processFiles() {
    this.state.successCount = 0;
    this.state.errorCount = 0;

    for (let i = 0; i < this.state.files.length; i++) {
      const file = this.state.files[i];
      try {
        const data = await readExcelFile(file);
        const processedData = ExcelProcessor.processSheet(data);

        this.displayResult(file.name, processedData);
        this.state.successCount++;
      } catch (error) {
        this.state.errorCount++;
        this.showError(`Erro ao processar ${file.name}: ${error.message}`);
      }

      this.updateProgress(((i + 1) / this.state.files.length) * 100);
    }

    this.showSummary();
  },

  // Exibe resultado processado
  displayResult(fileName, data) {
    console.log("Exibindo resultado para:", fileName, data); // <-- Adicionado

    const resultHtml = `
        <div class="card mb-3">
            <div class="card-header">
                <h5 class="card-title m-0">CRQ: ${fileName.replace(
                  /\.[^/.]+$/,
                  ""
                )}</h5>
            </div>
            <div class="card-body">
                <div class="row">
                    <div class="col-12">
                        <p><strong>Hora de Início:</strong> ${
                          data.horaInicio
                        } - ${data.dataInicio}</p>
                        <p><strong>Hora de Término:</strong> ${
                          data.horaTermino
                        } - ${data.dataTermino}</p>
                        <p><strong>Descrição da atividade:</strong>*${
                          data.beneficios
                        }*</p>
                        <p><strong>Área Afetada:</strong> *${
                          data.areaAfetada
                        }*</p>
                        <p><strong>Impactos:</strong> *${data.impactos}*</p>
                        <p><strong>Responsável:</strong> ${data.responsavel}</p>
                        <p>-----------------------------------------</p>
                    </div>
                </div>
            </div>
        </div>
    `;

    this.elements.resultsContainer.insertAdjacentHTML("beforeend", resultHtml);
  },

  // Atualiza lista de arquivos
  updateFileList() {
    if (!this.elements.fileList) return;

    const fileListHtml = this.state.files
      .map(
        (file) => `
            <div class="file-item">
                <i class="fas fa-file-excel"></i>
                <span>${file.name}</span>
            </div>
        `
      )
      .join("");

    this.elements.fileList.innerHTML = fileListHtml;
  },

  // Mostra/esconde indicador de carregamento
  showLoading(show) {
    if (this.elements.loadingIndicator) {
      this.elements.loadingIndicator.classList.toggle("hidden", !show);
    }
  },

  // Atualiza barra de progresso
  updateProgress(percentage) {
    if (this.elements.progressBar) {
      this.elements.progressBar.style.width = `${percentage}%`;
      this.elements.progressBar.setAttribute("aria-valuenow", percentage);
    }
  },

  // Limpa resultados anteriores
  clearResults() {
    if (this.elements.resultsContainer) {
      this.elements.resultsContainer.innerHTML = "";
    }
  },

  // Exibe notificação
  showNotification(type, message) {
    const toast = document.createElement("div");
    toast.className = `toast ${type} show`;
    toast.innerHTML = `
            <div class="toast-header">
                <strong class="me-auto">${
                  type === "success" ? "Sucesso" : "Erro"
                }</strong>
                <button type="button" class="btn-close" data-bs-dismiss="toast"></button>
            </div>
            <div class="toast-body">${message}</div>
        `;

    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), 5000);
  },

  // Exibe erro
  showError(message) {
    const errorHtml = `
            <div class="alert alert-danger alert-dismissible fade show" role="alert">
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            </div>
        `;

    this.elements.resultsContainer.insertAdjacentHTML("beforeend", errorHtml);
  },

  // Exibe resumo do processamento
  showSummary() {
    const summaryHtml = `
            <div class="alert alert-info">
                <h4>Resumo do Processamento</h4>
                <p>Arquivos processados com sucesso: ${this.state.successCount}</p>
                <p>Arquivos com erro: ${this.state.errorCount}</p>
            </div>
        `;

    this.elements.resultsContainer.insertAdjacentHTML(
      "afterbegin",
      summaryHtml
    );
  },
};

// Inicialização
document.addEventListener("DOMContentLoaded", () => App.init());

export default App;

import { DateUtils } from "./utils/DateUtils.js";
import { DebugUtils } from "./utils/DebugUtils.js";
import { ExcelService } from "./services/ExcelService.js";
import { ChartService } from "./services/ChartService.js";
import { JobData } from "./models/JobData.js";
import { initChartResizer } from "./chartResizer.js";

class App {
  static init() {
    this.bindEvents();
  }

  static bindEvents() {
    document
      .getElementById("fileInput")
      .addEventListener("change", this.handleFileInput.bind(this));
    document
      .getElementById("processButton")
      .addEventListener("click", this.processFiles.bind(this));
    document
      .getElementById("clearButton")
      .addEventListener("click", this.clearAll.bind(this));
  }

  static async handleFileInput(event) {
    try {
      const files = event.target.files;
      if (!files || files.length === 0) {
        throw new Error("Nenhum arquivo selecionado");
      }

      console.log("Arquivos selecionados:", files.length);
      this.workbooks = await ExcelService.loadWorkbooks(files);
      DebugUtils.addDebugInfo(`${files.length} arquivos carregados`);
    } catch (error) {
      console.error("Erro no handleFileInput:", error);
      DebugUtils.logError(`Erro ao carregar arquivos: ${error.message}`);
    }
  }

  static async processFiles() {
    try {
      console.log("Iniciando processamento...");
      if (!this.workbooks || this.workbooks.length === 0) {
        throw new Error("Nenhum arquivo carregado");
      }
      console.log("Workbooks carregados:", this.workbooks.length);

      const jobData = new JobData();
      console.log("JobData criado");

      for (const workbook of this.workbooks) {
        console.log("Processando workbook...");
        const processedData = ExcelService.processWorkbook(workbook);
        console.log("Dados processados:", processedData);

        // Mesclar dados do processedData com jobData
        processedData.jobStats.forEach((value, key) => {
          if (jobData.jobStats.has(key)) {
            jobData.jobStats.set(key, jobData.jobStats.get(key) + value);
          } else {
            jobData.jobStats.set(key, value);
          }
        });

        processedData.jobDurations.forEach((durations, key) => {
          if (jobData.jobDurations.has(key)) {
            jobData.jobDurations.get(key).push(...durations);
          } else {
            jobData.jobDurations.set(key, [...durations]);
          }
        });
      }

      console.log("Criando gráficos...");
      this.createCharts(jobData);
      console.log("Gráficos criados");

      // Adicione após a inicialização dos gráficos
      initChartResizer();

      DebugUtils.addDebugInfo("Processamento concluído");
    } catch (error) {
      console.error("Erro detalhado:", error);
      DebugUtils.logError(`Erro ao processar arquivos: ${error.message}`);
    }
  }

  static createCharts(jobData) {
    // Gráfico de Overdue
    const overdueData = jobData.getOverdueStats();
    ChartService.createChart(
      "overdueChart",
      "Jobs em Overdue (>60min)",
      overdueData,
      {
        label: "Quantidade de Overdue",
        getValue: (d) => d.overdueCount,
        yAxisLabel: "Número de Ocorrências",
        tooltipCallback: (item) => [
          `Overdue: ${item.overdueCount}`,
          `Total Execuções: ${item.totalExecutions}`,
          `Percentual: ${item.overduePercentage.toFixed(1)}%`,
          `Duração Média: ${DateUtils.formatDuration(item.avgDuration)}`,
        ],
      }
    );

    // Gráfico de Duração Média
    const durationData = jobData.getDurationStats();
    ChartService.createChart(
      "durationChart",
      "Duração Média dos Jobs",
      durationData,
      {
        label: "Duração Média",
        getValue: (d) => d.avg,
        yAxisLabel: "Duração (minutos)",
        tooltipCallback: (item) => [
          `Média: ${DateUtils.formatDuration(item.avg)}`,
          `Mín: ${DateUtils.formatDuration(item.min)}`,
          `Máx: ${DateUtils.formatDuration(item.max)}`,
          `Execuções: ${item.executionCount}`,
        ],
      }
    );

    // Gráfico de Variação
    const variationData = jobData.getVariationStats();
    ChartService.createChart(
      "variationChart",
      "Variação de Duração dos Jobs",
      variationData,
      {
        label: "Coeficiente de Variação",
        getValue: (d) => d.cv,
        yAxisLabel: "Variação (%)",
        tooltipCallback: (item) => [
          `CV: ${item.cv.toFixed(2)}%`,
          `Média: ${DateUtils.formatDuration(item.avgDuration)}`,
          `Desvio Padrão: ${DateUtils.formatDuration(item.stdDev)}`,
          `Execuções: ${item.executionCount}`,
          `Min: ${DateUtils.formatDuration(item.minDuration)}`,
          `Max: ${DateUtils.formatDuration(item.maxDuration)}`,
        ],
      }
    );
  }

  static clearAll() {
    ChartService.clearCharts();
    DebugUtils.clearDebugInfo();
    document.getElementById("errorLog").innerHTML = "";
    document.getElementById("fileInput").value = "";
  }
}

// Inicializar aplicação
document.addEventListener("DOMContentLoaded", () => App.init());

import { DateUtils } from "../utils/DateUtils.js";
import { JobData } from "../models/JobData.js";

export class ExcelService {
  static async loadWorkbooks(files) {
    try {
      console.log("Iniciando carregamento de arquivos:", files.length);
      const workbooks = [];

      for (const file of files) {
        try {
          console.log("Processando arquivo:", file.name);
          const data = await file.arrayBuffer();
          const workbook = XLSX.read(data);
          workbooks.push(workbook);
          console.log("Arquivo processado com sucesso:", file.name);
        } catch (error) {
          console.error(`Erro ao processar arquivo ${file.name}:`, error);
        }
      }

      console.log("Total de workbooks carregados:", workbooks.length);
      return workbooks;
    } catch (error) {
      console.error("Erro ao carregar workbooks:", error);
      throw error;
    }
  }

  static processWorkbook(workbook) {
    try {
      console.log("Iniciando processamento do workbook");
      const jobData = new JobData();
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      const data = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        dateNF: "mm/dd/yyyy hh:mm:ss AM/PM",
        defval: null,
      });

      const headerRow = 8; // linha 9 (índice 8)
      const headers = data[headerRow];

      let processed = 0;
      let errors = 0;
      let skipped = 0;

      for (let i = headerRow + 1; i < data.length; i++) {
        try {
          const row = data[i];
          if (!row || !row[0]) {
            skipped++;
            continue;
          }

          const jobName = String(row[0]).trim();
          const startDateStr = row[14]; // Coluna O
          const endDateStr = row[16]; // Coluna Q
          const durationStr = row[11]; // Coluna L

          // Log detalhado para depuração
          console.log(`Processando linha ${i + 1}:`, {
            jobName,
            duration: durationStr,
            startDate: startDateStr,
            endDate: endDateStr,
          });

          const startDate = DateUtils.parseDateTime(startDateStr);
          const endDate = DateUtils.parseDateTime(endDateStr);
          let duration = DateUtils.parseDuration(durationStr);

          // Log dos valores parseados
          console.log(`Valores parseados linha ${i + 1}:`, {
            startDate: startDate ? startDate.toISOString() : null,
            endDate: endDate ? endDate.toISOString() : null,
            duration,
          });

          // Se não tiver duração válida, calcular das datas
          if (duration === 0 && startDate && endDate) {
            duration = DateUtils.calculateDurationBetweenDates(
              startDate,
              endDate
            );
            console.log(`Duração calculada das datas: ${duration} minutos`);
          }

          if (duration > 0) {
            jobData.addJob(jobName, duration, startDate, endDate);
            processed++;
            console.log(
              `Job processado com sucesso: ${jobName}, duração: ${duration} minutos`
            );
          } else {
            console.warn(
              `Linha ${i + 1}: Não foi possível determinar duração`,
              {
                jobName,
                startDateStr,
                endDateStr,
                durationStr,
                calculatedDuration: duration,
              }
            );
            errors++;
          }
        } catch (error) {
          errors++;
          console.error(`Erro na linha ${i + 1}:`, error);
        }
      }

      console.log("Resumo do processamento:", {
        totalLinhas: data.length - headerRow - 1,
        processadas: processed,
        erros: errors,
        ignoradas: skipped,
        jobsUnicos: jobData.jobDurations.size,
      });

      return jobData;
    } catch (error) {
      console.error("Erro ao processar workbook:", error);
      throw error;
    }
  }

  static findHeaderRow(data) {
    return data.findIndex(
      (row) =>
        row &&
        row.some(
          (cell) => cell && String(cell).toLowerCase().includes("start date")
        )
    );
  }

  static getColumnIndices(headers) {
    console.log("Analisando cabeçalhos:", headers);

    // Procura pelas colunas específicas
    const indices = {
      jobName: 0, // Coluna A (fixo)
      startDateTime: -1,
      endDateTime: -1,
      duration: -1,
    };

    headers.forEach((header, index) => {
      if (!header) return;
      const headerStr = String(header).toLowerCase();

      if (headerStr.includes("start date")) {
        indices.startDateTime = index;
      } else if (headerStr.includes("end date")) {
        indices.endDateTime = index;
      } else if (headerStr.includes("duration")) {
        indices.duration = index;
      }
    });

    console.log("Índices encontrados:", indices);
    return indices;
  }

  static processWorkbook(workbook) {
    try {
      const jobData = new JobData();
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      // Converter para array mantendo valores originais
      const data = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: true,
        dateNF: "mm/dd/yyyy hh:mm:ss AM/PM",
      });

      console.log("Primeiras linhas:", data.slice(0, 5));

      // Encontrar linha do cabeçalho (normalmente linha 9)
      const headerRow = 8; // fixo na linha 9 (índice 8)
      const headers = data[headerRow];
      const indices = this.getColumnIndices(headers);

      let processed = 0;
      let errors = 0;

      for (let i = headerRow + 1; i < data.length; i++) {
        try {
          const row = data[i];
          if (!row || !row[indices.jobName]) continue;

          const jobName = String(row[indices.jobName]).trim();
          let startDate = null;
          let endDate = null;
          let duration = 0;

          // Tentar obter as datas
          if (indices.startDateTime !== -1 && indices.endDateTime !== -1) {
            startDate = DateUtils.parseDateTime(row[indices.startDateTime]);
            endDate = DateUtils.parseDateTime(row[indices.endDateTime]);

            if (startDate && endDate) {
              duration = DateUtils.calculateDurationBetweenDates(
                startDate,
                endDate
              );
            }
          }

          // Se tiver uma duração específica na planilha, usar ela
          if (indices.duration !== -1 && row[indices.duration]) {
            const durationFromCell = DateUtils.parseDuration(
              row[indices.duration]
            );
            if (durationFromCell > 0) {
              duration = durationFromCell;
            }
          }

          if (duration > 0) {
            jobData.addJob(jobName, duration, startDate, endDate);
            processed++;
          } else {
            errors++;
            console.warn(
              `Linha ${
                i + 1
              }: Não foi possível determinar duração para ${jobName}`
            );
          }
        } catch (error) {
          errors++;
          console.error(`Erro na linha ${i + 1}:`, error);
        }
      }

      console.log(
        `Processamento concluído: ${processed} linhas processadas, ${errors} erros`
      );
      return jobData;
    } catch (error) {
      console.error("Erro ao processar workbook:", error);
      throw error;
    }
  }

  static findHeaderRow(data) {
    // Procurar especificamente na linha 9 (índice 8)
    const headerRow = 8;

    // Validar se a linha existe e contém os cabeçalhos esperados
    if (
      data[headerRow] &&
      data[headerRow][14]?.toLowerCase().includes("start date")
    ) {
      // Coluna O
      console.log("Cabeçalhos encontrados na linha 9");
      return headerRow;
    }

    console.warn("Cabeçalhos não encontrados na linha 9");
    return -1;
  }

  static processWorkbook(workbook) {
    try {
      console.log("Iniciando processamento do workbook");
      const jobData = new JobData();
      const sheet = workbook.Sheets[workbook.SheetNames[0]];

      console.log("Informações da planilha:", {
        nome: workbook.SheetNames[0],
        range: sheet["!ref"],
      });

      const data = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        raw: false,
        dateNF: "mm/dd/yyyy hh:mm:ss AM/PM",
        defval: null,
      });

      // Verificar se temos pelo menos 9 linhas
      if (data.length < 9) {
        throw new Error("Planilha não contém linhas suficientes");
      }

      const headerRow = this.findHeaderRow(data);
      if (headerRow === -1) {
        throw new Error("Cabeçalhos não encontrados na linha 9");
      }

      const headers = data[headerRow];
      const indices = this.getColumnIndices(headers);

      // Log das primeiras linhas após o cabeçalho
      console.log("Exemplo de dados:", {
        linha10: data[headerRow + 1],
        linha11: data[headerRow + 2],
      });

      let processed = 0;
      let errors = 0;

      for (let i = headerRow + 1; i < data.length; i++) {
        try {
          const row = data[i];
          if (!row || !row[indices.jobName]) continue;

          const jobName = String(row[indices.jobName]).trim();
          if (!jobName) continue;

          // Log detalhado das primeiras linhas processadas
          if (processed < 3) {
            console.log(`Processando linha ${i + 1}:`, {
              jobName,
              startDate: row[indices.startDateTime],
              endDate: row[indices.endDateTime],
              duration: row[indices.duration],
            });
          }

          let duration =
            indices.duration !== -1
              ? DateUtils.parseDuration(row[indices.duration])
              : 0;

          if (duration === 0) {
            const startDate = DateUtils.parseDateTime(
              row[indices.startDateTime]
            );
            const endDate = DateUtils.parseDateTime(row[indices.endDateTime]);

            if (startDate && endDate) {
              duration = DateUtils.calculateDurationBetweenDates(
                startDate,
                endDate
              );
            }
          }

          if (duration > 0) {
            jobData.addJob(jobName, duration);
            processed++;
          } else {
            errors++;
          }
        } catch (error) {
          errors++;
          console.error(`Erro na linha ${i + 1}:`, error);
        }
      }

      console.log("Resumo do processamento:", {
        totalLinhas: data.length - headerRow - 1,
        processadas: processed,
        erros: errors,
      });

      return jobData;
    } catch (error) {
      console.error("Erro ao processar workbook:", error);
      throw error;
    }
  }

  static findColumnByName(headers, searchName) {
    return headers.findIndex(
      (header) =>
        header &&
        String(header).toLowerCase().includes(searchName.toLowerCase())
    );
  }
}

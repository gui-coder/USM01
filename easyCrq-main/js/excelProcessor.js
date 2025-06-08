const ExcelProcessor = {
  processSheet(data) {
    try {
      const workbook = XLSX.read(data, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];

      const locations = this.findHeaderLocations(worksheet);
      console.log("Localizações de cabeçalho encontradas:", locations);
      console.log("Worksheet ref:", worksheet["!ref"]);

      const horarios = this.extractTimes(worksheet, locations);
      const datas = this.extractDatesFromColumns(worksheet, locations);
      console.log("Datas extraídas:", datas);

      return {
        horaInicio: horarios.inicio,
        horaTermino: horarios.termino,
        dataInicio: datas.dataInicio,
        dataTermino: datas.dataTermino,
        beneficios: this.getValueFromNextCell(
          worksheet,
          locations.beneficios,
          "beneficios"
        ),
        impactos: this.getValueFromNextCell(
          worksheet,
          locations.impactos,
          "default"
        ),
        areaAfetada: this.getValueFromNextCell(
          worksheet,
          locations.areaAfetada,
          "default"
        ),
        responsavel: this.getValueFromNextCell(
          worksheet,
          locations.responsavel,
          "responsavel"
        ),
      };
    } catch (error) {
      console.error("Erro ao processar planilha:", error);
      throw new Error("Falha ao processar planilha");
    }
  },

  findHeaderLocations(worksheet) {
    const headers = {
      horaInicio: ["início", "inicio"],
      horaTermino: ["término", "termino"],
      beneficios: ["escopo da manutenção:"],
      impactos: ["impactos", "impactos da manutenção"],
      areaAfetada: ["área afetada", "area afetada", "empresa afetada"],
      responsavel: ["responsável", "responsavel", "analista responsável"],
    };

    const locations = {};
    const range = XLSX.utils.decode_range(worksheet["!ref"]);

    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[address];

        if (!cell || !cell.v) continue;

        const value = cell.v.toString().toLowerCase().trim();

        for (const [key, terms] of Object.entries(headers)) {
          if (!locations[key] && terms.some((term) => value.includes(term))) {
            locations[key] = { row, col, address };
            console.log(
              `Encontrado cabeçalho ${key} em ${address} com valor "${value}"`
            );
            break;
          }
        }
      }
    }

    return locations;
  },

  extractTimes(worksheet, locations) {
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    let horasInicio = [];
    let horasTermino = [];

    // Aceita números Excel, "hh:mm", "hh:mm:ss", "h:mm", "h:mm:ss", "hhhmm", "hhh:mm"
    const isTimeValue = (value) => {
      if (typeof value === "number" && value > 0 && value < 1) return true;
      if (
        typeof value === "string" &&
        value.trim() &&
        (/^(\d{1,2}):(\d{2})(:\d{2})?$/.test(value.trim()) || // 22:00, 7:30, 22:00:00
          /^(\d{1,2})h(\d{2})$/.test(value.trim().toLowerCase()) || // 22h00, 7h30
          /^(\d{1,2})h:(\d{2})$/.test(value.trim().toLowerCase())) // 22h:00
      )
        return true;
      return false;
    };

    // Normaliza para "hh:mm"
    const normalizeTimeString = (val) => {
      if (typeof val === "string") {
        let v = val.trim().toLowerCase();
        let match;
        if ((match = v.match(/^(\d{1,2}):(\d{2})(:\d{2})?$/))) {
          return `${match[1].padStart(2, "0")}:${match[2]}`;
        }
        if ((match = v.match(/^(\d{1,2})h(\d{2})$/))) {
          return `${match[1].padStart(2, "0")}:${match[2]}`;
        }
        if ((match = v.match(/^(\d{1,2})h:(\d{2})$/))) {
          return `${match[1].padStart(2, "0")}:${match[2]}`;
        }
      }
      return val;
    };

    // Busca horários de início (até 20 linhas abaixo do cabeçalho)
    if (locations.horaInicio) {
      const maxBusca = Math.min(range.e.r, locations.horaInicio.row + 20);
      for (let row = locations.horaInicio.row; row <= maxBusca; row++) {
        const cell =
          worksheet[
            XLSX.utils.encode_cell({ r: row, c: locations.horaInicio.col })
          ];
        if (cell && isTimeValue(cell.v)) {
          let val = cell.v;
          if (typeof val === "string") val = normalizeTimeString(val);
          horasInicio.push(val);
        }
      }
    }

    // Busca horários de término (até 20 linhas abaixo do cabeçalho)
    if (locations.horaTermino) {
      const maxBusca = Math.min(range.e.r, locations.horaTermino.row + 20);
      for (let row = locations.horaTermino.row; row <= maxBusca; row++) {
        const cell =
          worksheet[
            XLSX.utils.encode_cell({ r: row, c: locations.horaTermino.col })
          ];
        if (cell && isTimeValue(cell.v)) {
          let val = cell.v;
          if (typeof val === "string") val = normalizeTimeString(val);
          horasTermino.push(val);
        }
      }
    }

    // Filtra valores inválidos
    horasInicio = horasInicio.filter((v) => v && v !== "-" && v !== "N/A");
    horasTermino = horasTermino.filter((v) => v && v !== "-" && v !== "N/A");

    // Pega o menor valor para início e o maior para término
    let inicio = "";
    let termino = "";

    const onlyNumbers = (arr) => arr.filter((v) => typeof v === "number");
    const onlyStrings = (arr) => arr.filter((v) => typeof v === "string");

    if (horasInicio.length > 0) {
      if (onlyNumbers(horasInicio).length > 0) {
        inicio = this.formatTime(Math.min(...onlyNumbers(horasInicio)));
      } else {
        inicio = onlyStrings(horasInicio).sort()[0];
      }
    }

    if (horasTermino.length > 0) {
      if (onlyNumbers(horasTermino).length > 0) {
        termino = this.formatTime(Math.max(...onlyNumbers(horasTermino)));
      } else {
        termino = onlyStrings(horasTermino).sort().reverse()[0];
      }
    }

    return { inicio, termino };
  },

  formatTime(value) {
    if (!value) return "";

    const totalHoras = value * 24;
    const horas = Math.floor(totalHoras);
    const minutos = Math.floor((totalHoras - horas) * 60);

    return `${horas.toString().padStart(2, "0")}:${minutos
      .toString()
      .padStart(2, "0")}`;
  },

  getValueFromNextCell(worksheet, location, type = "default") {
    if (!location) {
      console.log(`Nenhuma localização encontrada para ${type}`);
      return "";
    }

    const rightCell =
      worksheet[
        XLSX.utils.encode_cell({
          r: location.row,
          c: location.col + 1,
        })
      ];
    const bottomCell =
      worksheet[
        XLSX.utils.encode_cell({
          r: location.row + 1,
          c: location.col,
        })
      ];

    const rightValue = rightCell && rightCell.v ? rightCell.v.toString() : null;
    const bottomValue =
      bottomCell && bottomCell.v ? bottomCell.v.toString() : null;

    console.log(`Valor à direita de ${type}:`, rightValue);
    console.log(`Valor abaixo de ${type}:`, bottomValue);

    if (type === "responsavel") {
      return bottomValue || "";
    }

    return rightValue || bottomValue || "";
  },

  extractDate(worksheet) {
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    const possibleDatePatterns = [
      /^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/, // 2/07/2025, 02-07-25, 2-7-2025
      /^\d{1,2}[-\/][a-z]{3,}$/, // 2-nov, 12-jul
      /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/, // 2025-07-02
      /^\d{1,2}[-\/]\d{1,2}$/, // 2/07, 12-07
    ];

    // Função para normalizar para dd/mm/yyyy
    const normalizeDate = (val) => {
      if (typeof val === "number") {
        const excelEpoch = new Date(1899, 11, 30);
        const utc = excelEpoch.getTime() + val * 24 * 60 * 60 * 1000;
        const date = new Date(utc);
        const day = date.getDate().toString().padStart(2, "0");
        const month = (date.getMonth() + 1).toString().padStart(2, "0");
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
      }

      if (typeof val !== "string") return null;
      let v = val.trim().toLowerCase();
      let match;

      // 2-nov ou 12-jul
      match = v.match(/^(\d{1,2})[-\/]([a-z]{3,})$/);
      if (match) {
        const day = match[1].padStart(2, "0");
        const monthStr = match[2];
        const months = {
          jan: "01",
          fev: "02",
          feb: "02",
          mar: "03",
          abr: "04",
          apr: "04",
          mai: "05",
          may: "05",
          jun: "06",
          jul: "07",
          ago: "08",
          aug: "08",
          set: "09",
          sep: "09",
          out: "10",
          oct: "10",
          nov: "11",
          dez: "12",
          dec: "12",
        };
        const month = months[monthStr.slice(0, 3)] || "01";
        const year = new Date().getFullYear();
        return `${day}/${month}/${year}`;
      }

      // 2/07/2025 ou 02-07-25 ou 2-7-2025
      match = v.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
      if (match) {
        const day = match[1].padStart(2, "0");
        const month = match[2].padStart(2, "0");
        let year = match[3];
        if (year.length === 2) year = "20" + year;
        return `${day}/${month}/${year}`;
      }

      // 2025-07-02
      match = v.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
      if (match) {
        const year = match[1];
        const month = match[2].padStart(2, "0");
        const day = match[3].padStart(2, "0");
        return `${day}/${month}/${year}`;
      }

      // 2/07 ou 12-07 (sem ano, assume ano atual)
      match = v.match(/^(\d{1,2})[-\/](\d{1,2})$/);
      if (match) {
        const day = match[1].padStart(2, "0");
        const month = match[2].padStart(2, "0");
        const year = new Date().getFullYear();
        return `${day}/${month}/${year}`;
      }

      return null;
    };

    // Busca por data em toda a planilha (até 30 linhas e 10 colunas)
    for (
      let row = range.s.r;
      row <= Math.min(range.e.r, range.s.r + 30);
      row++
    ) {
      for (
        let col = range.s.c;
        col <= Math.min(range.e.c, range.s.c + 10);
        col++
      ) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[address];
        if (!cell || !cell.v) continue;
        const value = cell.v; // NÃO use .toString().trim() aqui!
        console.log(
          "Valor lido na busca geral (linha",
          row,
          "coluna",
          col,
          "):",
          value
        ); // <-- ADICIONE
        for (const pattern of possibleDatePatterns) {
          if (pattern.test(value.toLowerCase())) {
            const normalized = normalizeDate(value);
            if (normalized) return normalized;
          }
        }
      }
    }
    return "";
  },

  extractDatesFromColumns(worksheet, locations) {
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    const possibleDatePatterns = [
      /^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/, // 2/07/2025, 02-07-25, 2-7-2025
      /^\d{1,2}[-\/][a-z]{3,}$/, // 2-nov, 12-jul
      /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/, // 2025-07-02
      /^\d{1,2}[-\/]\d{1,2}$/, // 2/07, 12-07
    ];

    // Normaliza para dd/mm/yyyy
    const normalizeDate = (val) => {
      if (typeof val === "number") {
        const excelEpoch = new Date(1899, 11, 30);
        const utc = excelEpoch.getTime() + val * 24 * 60 * 60 * 1000;
        const date = new Date(utc);
        const day = date.getDate().toString().padStart(2, "0");
        const month = (date.getMonth() + 1).toString().padStart(2, "0");
        const year = date.getFullYear();
        return `${day}/${month}/${year}`;
      }

      if (typeof val !== "string") return null;
      let v = val.trim().toLowerCase();
      let match;

      // 2-nov ou 12-jul
      match = v.match(/^(\d{1,2})[-\/]([a-z]{3,})$/);
      if (match) {
        const day = match[1].padStart(2, "0");
        const monthStr = match[2];
        const months = {
          jan: "01",
          fev: "02",
          feb: "02",
          mar: "03",
          abr: "04",
          apr: "04",
          mai: "05",
          may: "05",
          jun: "06",
          jul: "07",
          ago: "08",
          aug: "08",
          set: "09",
          sep: "09",
          out: "10",
          oct: "10",
          nov: "11",
          dez: "12",
          dec: "12",
        };
        const month = months[monthStr.slice(0, 3)] || "01";
        const year = new Date().getFullYear();
        return `${day}/${month}/${year}`;
      }

      // 2/07/2025 ou 02-07-25 ou 2-7-2025
      match = v.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{2,4})$/);
      if (match) {
        const day = match[1].padStart(2, "0");
        const month = match[2].padStart(2, "0");
        let year = match[3];
        if (year.length === 2) year = "20" + year;
        return `${day}/${month}/${year}`;
      }

      // 2025-07-02
      match = v.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
      if (match) {
        const year = match[1];
        const month = match[2].padStart(2, "0");
        const day = match[3].padStart(2, "0");
        return `${day}/${month}/${year}`;
      }

      // 2/07 ou 12-07 (sem ano, assume ano atual)
      match = v.match(/^(\d{1,2})[-\/](\d{1,2})$/);
      if (match) {
        const day = match[1].padStart(2, "0");
        const month = match[2].padStart(2, "0");
        const year = new Date().getFullYear();
        return `${day}/${month}/${year}`;
      }

      return null;
    };

    // 1. Procura por cabeçalhos de data em toda a planilha
    const headerTerms = [
      "data",
      "data início",
      "data inicio",
      "data fim",
      "data término",
      "data termino",
    ];
    let dataCols = [];
    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const address = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = worksheet[address];
        if (!cell || !cell.v) continue;
        const value = cell.v.toString().toLowerCase().trim();
        if (headerTerms.some((term) => value.includes(term))) {
          dataCols.push({ col, row });
        }
      }
    }

    // 2. Para cada coluna de data encontrada, busca datas nas linhas seguintes
    let datasEncontradas = [];
    dataCols.forEach(({ col, row }) => {
      for (let r = row + 1; r <= Math.min(range.e.r, row + 20); r++) {
        const address = XLSX.utils.encode_cell({ r, c: col });
        const cell = worksheet[address];
        if (!cell || !cell.v) continue;
        const value = cell.v; // NÃO use .toString().trim() aqui!
        console.log(
          `Valor lido na coluna de data (coluna ${col}, linha ${r}):`,
          value
        );
        let normalized = null;
        if (typeof value === "number") {
          normalized = normalizeDate(value);
        } else if (typeof value === "string") {
          for (const pattern of possibleDatePatterns) {
            if (pattern.test(value.toLowerCase())) {
              normalized = normalizeDate(value);
              break;
            }
          }
        }
        if (normalized) {
          console.log(
            "Data reconhecida e normalizada:",
            value,
            "->",
            normalized
          );
          datasEncontradas.push(normalized);
        }
      }
    });

    // 3. Se não encontrou nada, faz busca geral por datas (fallback)
    if (datasEncontradas.length === 0) {
      for (
        let row = range.s.r;
        row <= Math.min(range.e.r, range.s.r + 30);
        row++
      ) {
        for (
          let col = range.s.c;
          col <= Math.min(range.e.c, range.s.c + 10);
          col++
        ) {
          const address = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[address];
          if (!cell || !cell.v) continue;
          const value = cell.v; // NÃO use .toString().trim() aqui!
          console.log(
            `Valor lido na busca geral (coluna ${col}, linha ${row}):`,
            value
          );
          let normalized = null;
          if (typeof value === "number") {
            normalized = normalizeDate(value);
          } else if (typeof value === "string") {
            for (const pattern of possibleDatePatterns) {
              if (pattern.test(value.toLowerCase())) {
                normalized = normalizeDate(value);
                break;
              }
            }
          }
          if (normalized) {
            console.log(
              "Data reconhecida e normalizada (fallback):",
              value,
              "->",
              normalized
            );
            datasEncontradas.push(normalized);
          }
        }
      }
    }

    // 4. Decide início e fim
    const toDateObj = (str) => {
      if (!str || typeof str !== "string" || !str.includes("/")) return null;
      const [d, m, y] = str.split("/");
      return new Date(`${y}-${m}-${d}`);
    };

    let dataInicio = "";
    let dataTermino = "";

    if (datasEncontradas.length === 1) {
      dataInicio = dataTermino = datasEncontradas[0];
    } else if (datasEncontradas.length > 1) {
      const datasObj = datasEncontradas
        .map(toDateObj)
        .filter(Boolean)
        .sort((a, b) => a - b);
      dataInicio = datasObj[0]
        ? `${datasObj[0].getDate().toString().padStart(2, "0")}/${(
            datasObj[0].getMonth() + 1
          )
            .toString()
            .padStart(2, "0")}/${datasObj[0].getFullYear()}`
        : "";
      dataTermino = datasObj[datasObj.length - 1]
        ? `${datasObj[datasObj.length - 1]
            .getDate()
            .toString()
            .padStart(2, "0")}/${(datasObj[datasObj.length - 1].getMonth() + 1)
            .toString()
            .padStart(2, "0")}/${datasObj[datasObj.length - 1].getFullYear()}`
        : "";
    }

    return { dataInicio, dataTermino };
  },
};

export default ExcelProcessor;

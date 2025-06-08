// DateUtils.js
export class DateUtils {
  static parseDuration(value) {
    if (!value) return 0;

    try {
      // Se já for um número
      if (typeof value === "number") {
        return value;
      }

      const str = String(value).trim();

      // Formato "X' Y''" (minutos e segundos)
      const timeMatch = str.match(/(\d+)'\s*(\d+)''/);
      if (timeMatch) {
        const minutes = parseInt(timeMatch[1], 10) || 0;
        const seconds = parseInt(timeMatch[2], 10) || 0;
        return minutes + seconds / 60;
      }

      // Ignorar se for um nome de job
      if (str.match(/^[A-Z0-9_-]+$/)) {
        return 0;
      }

      // Tentar converter diretamente para número
      const num = parseFloat(str);
      if (!isNaN(num)) {
        return num;
      }

      return 0;
    } catch (error) {
      console.error("Erro ao processar duração:", error);
      return 0;
    }
  }

  static parseDateTime(value) {
    if (!value) return null;

    try {
      // Se já for um Date
      if (value instanceof Date) {
        return value;
      }

      // Se for número (formato Excel)
      if (typeof value === "number") {
        // Converter número do Excel para data
        return new Date((value - 25569) * 86400 * 1000);
      }

      // Se for string
      const str = String(value).trim();
      const date = new Date(str);

      if (!isNaN(date.getTime())) {
        return date;
      }

      return null;
    } catch (error) {
      console.error("Erro ao processar data:", error);
      return null;
    }
  }

  static calculateDurationBetweenDates(startDate, endDate) {
    if (!startDate || !endDate) return 0;
    if (!(startDate instanceof Date) || !(endDate instanceof Date)) return 0;

    const diffMs = endDate.getTime() - startDate.getTime();
    const minutes = diffMs / (1000 * 60);

    if (minutes < 0) {
      console.warn("Duração negativa detectada:", { startDate, endDate });
      return 0;
    }

    return minutes;
  }
  static formatDuration(duration) {
    if (!duration && duration !== 0) return "N/A";

    const minutes = typeof duration === "number" ? duration : 0;

    if (minutes < 1) {
      return `${Math.round(minutes * 60)}s`;
    } else if (minutes < 60) {
      return `${Math.round(minutes)}min`;
    } else {
      const hours = Math.floor(minutes / 60);
      const remainingMinutes = Math.round(minutes % 60);
      return `${hours}h ${remainingMinutes}min`;
    }
  }
}

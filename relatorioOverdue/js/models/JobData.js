export class JobData {
  constructor() {
    this.jobStats = new Map(); // Para contar overdues
    this.jobDurations = new Map(); // Para armazenar todas as durações
    this.jobExecutions = new Map(); // Para armazenar detalhes de cada execução
  }

  addJob(jobName, duration, startDate = null, endDate = null) {
    if (!jobName || duration <= 0) {
      console.warn("Dados inválidos:", { jobName, duration });
      return;
    }

    const normalizedName = String(jobName).trim();

    // Armazenar detalhes da execução
    if (!this.jobExecutions.has(normalizedName)) {
      this.jobExecutions.set(normalizedName, []);
    }

    this.jobExecutions.get(normalizedName).push({
      duration,
      startDate,
      endDate,
      timestamp: new Date(),
    });

    // Armazenar duração para análise estatística
    if (!this.jobDurations.has(normalizedName)) {
      this.jobDurations.set(normalizedName, []);
    }
    this.jobDurations.get(normalizedName).push(duration);

    // Contabilizar overdue se duração > 60 minutos
    if (duration > 60) {
      this.jobStats.set(
        normalizedName,
        (this.jobStats.get(normalizedName) || 0) + 1
      );
    }
  }

  getJobAnalysis(jobName) {
    const durations = this.jobDurations.get(jobName) || [];
    const executions = this.jobExecutions.get(jobName) || [];

    if (durations.length === 0) return null;

    const avg = durations.reduce((a, b) => a + b, 0) / durations.length;
    const variance =
      durations.reduce((a, b) => a + Math.pow(b - avg, 2), 0) /
      durations.length;
    const stdDev = Math.sqrt(variance);

    return {
      totalExecutions: durations.length,
      overdueCount: this.jobStats.get(jobName) || 0,
      durations: {
        min: Math.min(...durations),
        max: Math.max(...durations),
        avg: avg,
        stdDev: stdDev,
        cv: (stdDev / avg) * 100, // Coeficiente de variação
      },
      executionTimes: executions.map((e) => ({
        duration: e.duration,
        startDate: e.startDate,
        endDate: e.endDate,
      })),
    };
  }

  getOverdueStats() {
    return Array.from(this.jobStats.entries())
      .map(([name, count]) => {
        const analysis = this.getJobAnalysis(name);
        return {
          name,
          overdueCount: count,
          totalExecutions: analysis.totalExecutions,
          overduePercentage: (count / analysis.totalExecutions) * 100,
          avgDuration: analysis.durations.avg,
        };
      })
      .sort((a, b) => b.overdueCount - a.overdueCount);
  }

  getDurationStats() {
    return Array.from(this.jobDurations.entries())
      .map(([name, _]) => {
        const analysis = this.getJobAnalysis(name);
        return {
          name,
          ...analysis.durations,
          executionCount: analysis.totalExecutions,
        };
      })
      .sort((a, b) => b.avg - a.avg);
  }

  getVariationStats() {
    return Array.from(this.jobDurations.entries())
      .map(([name, _]) => {
        const analysis = this.getJobAnalysis(name);
        return {
          name,
          cv: analysis.durations.cv,
          avgDuration: analysis.durations.avg,
          stdDev: analysis.durations.stdDev,
          executionCount: analysis.totalExecutions,
          minDuration: analysis.durations.min,
          maxDuration: analysis.durations.max,
        };
      })
      .sort((a, b) => b.cv - a.cv);
  }

  getDetailedStats() {
    const stats = {
      totalJobs: this.jobDurations.size,
      totalExecutions: 0,
      jobsWithOverdue: 0,
      totalOverdues: 0,
      executionsPerJob: [],
    };

    this.jobDurations.forEach((durations, jobName) => {
      const analysis = this.getJobAnalysis(jobName);
      stats.totalExecutions += analysis.totalExecutions;

      if (analysis.overdueCount > 0) {
        stats.jobsWithOverdue++;
        stats.totalOverdues += analysis.overdueCount;
      }

      stats.executionsPerJob.push({
        jobName,
        executions: analysis.totalExecutions,
        avgDuration: analysis.durations.avg,
        variation: analysis.durations.cv,
      });
    });

    return stats;
  }
}

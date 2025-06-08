/**
 * Service for creating and managing Chart.js instances.
 *
 * @class ChartService
 */

/**
 * Creates a new chart instance on the specified canvas element.
 * If a previous chart exists on the given canvas, it will be destroyed before creating the new one.
 *
 * @static
 * @param {string} canvasId - The ID of the canvas element where the chart will be rendered.
 * @param {string} title - The title of the chart.
 * @param {Array<Object>} data - The array of data objects to be visualized; each object should contain at least a 'name' and 'value' property.
 * @param {Object} [options={}] - An optional object to customize the chart's appearance and behavior.
 * @param {string} [options.label] - Custom label for the dataset; defaults to the provided title if not specified.
 * @param {string} [options.type='bar'] - The type of chart to render (e.g., 'bar', 'line').
 * @param {string} [options.yAxisLabel] - The label for the Y-axis.
 * @param {Function} [options.getValue] - A function that extracts the numerical value from a data object; defaults to retrieving the 'value' property.
 * @param {Function} [options.formatY] - A callback function to format Y-axis tick values.
 * @param {Function} [options.tooltipCallback] - A custom callback for formatting tooltip labels.
 *
 * @returns {void}
 */

/**
 * Clears all managed charts by destroying existing chart instances and removing content from their containers.
 *
 * @static
 *
 * @returns {void}
 */

export class ChartService {
  static createChart(canvasId, title, data, options = {}) {
    const ctx = document.getElementById(canvasId);
    if (!ctx) return;

    // Destruir gráfico anterior se existir
    if (window[canvasId + "Chart"]) {
      window[canvasId + "Chart"].destroy();
    }

    const chartData = {
      labels: data.map((d) => d.name),
      datasets: [
        {
          label: options.label || title,
          data: data.map((d) =>
            options.getValue ? options.getValue(d) : d.value
          ),
          backgroundColor: "rgba(54, 162, 235, 0.5)",
          borderColor: "rgba(54, 162, 235, 1)",
          borderWidth: 1,
        },
      ],
    };

    window[canvasId + "Chart"] = new Chart(ctx, {
      type: options.type || "bar",
      data: chartData,
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: {
            beginAtZero: true,
            title: {
              display: true,
              text: options.yAxisLabel || "",
            },
            ticks: {
              callback: options.formatY || ((value) => value),
            },
          },
        },
        plugins: {
          tooltip: {
            callbacks: {
              label: function (context) {
                if (options.tooltipCallback) {
                  return options.tooltipCallback(data[context.dataIndex]);
                }
                return `${context.dataset.label}: ${context.formattedValue}`;
              },
            },
          },
          title: {
            display: true,
            text: title,
          },
        },
      },
    });
  }

  static clearCharts() {
    ["overdueChart", "durationChart", "variationChart"].forEach((id) => {
      const chart = window[id + "Chart"];
      if (chart) {
        chart.destroy();
        window[id + "Chart"] = null;
      }
      // Limpar containers de múltiplos gráficos
      const container = document.getElementById(`${id}-container`);
      if (container) {
        container.innerHTML = "";
      }
    });
  }
}

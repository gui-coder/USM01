export function initChartResizer() {
  const MIN_HEIGHT = 200;
  const MAX_HEIGHT = 800;
  const STEP = 100;

  document.querySelectorAll(".resize-btn").forEach((button) => {
    button.addEventListener("click", (e) => {
      const chartId = button.dataset.chart;
      const action = button.dataset.action;
      const chartContainer = document.querySelector(
        `#${chartId}`
      ).parentElement;

      let currentHeight = parseInt(getComputedStyle(chartContainer).height);

      if (action === "increase" && currentHeight < MAX_HEIGHT) {
        currentHeight += STEP;
      } else if (action === "decrease" && currentHeight > MIN_HEIGHT) {
        currentHeight -= STEP;
      }

      chartContainer.style.height = `${currentHeight}px`;

      // Força o Chart.js a se redimensionar
      const chartInstance = Chart.getChart(chartId);
      if (chartInstance) {
        chartInstance.resize();
      }
    });
  });
}

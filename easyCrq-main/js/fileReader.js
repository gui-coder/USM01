export function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new window.FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        resolve(data);
      } catch (error) {
        reject(new Error(`Erro ao ler arquivo: ${error.message}`));
      }
    };

    reader.onerror = () => reject(new Error("Erro ao ler arquivo"));
    reader.readAsArrayBuffer(file);
  });
}

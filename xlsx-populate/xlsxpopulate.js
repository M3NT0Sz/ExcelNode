const XlsxPopulate = require('xlsx-populate');
const fs = require('fs').promises;

// Caminho para o arquivo Excel
const excelFilePath = 'ModeloCreatorAuthor.xlsx';
// Caminho para o arquivo JSON de saída
const jsonFilePath = 'dadosxlsx.json';

// Função para ler o arquivo Excel e converter para JSON
async function excelToJson(excelFilePath, jsonFilePath) {
  try {
    // Carregar o arquivo Excel
    const workbook = await XlsxPopulate.fromFileAsync(excelFilePath);

    // Pegar a primeira planilha
    const sheet = workbook.sheet(0);

    // Obter os valores de todas as células
    const values = sheet.usedRange().value();

    // Extrair cabeçalhos e dados
    const headers = values[0];
    const data = values.slice(1);

    // Converter para JSON
    const jsonData = data.map(row => {
      const rowData = {};
      row.forEach((value, index) => {
        rowData[headers[index]] = value;
      });
      return rowData;
    });

    // Salvar os dados JSON em um arquivo
    await fs.writeFile(jsonFilePath, JSON.stringify(jsonData, null, 2));
    console.log('Dados do Excel convertidos para JSON e salvos em', jsonFilePath);
  } catch (error) {
    throw new Error('Erro ao converter Excel para JSON:', error);
  }
}

// Usar a função para converter Excel para JSON e salvar em um arquivo JSON
excelToJson(excelFilePath, jsonFilePath)
  .catch(error => {
    console.error(error.message);
  });

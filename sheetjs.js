const XLSX = require("xlsx");
const fs = require("fs").promises;

// Função para remover propriedades com valores null de um objeto
function removeNullProperties(obj) {
  return Object.fromEntries(Object.entries(obj).filter(([_, v]) => v !== null));
}

// Função para converter as opções em um único objeto "opcoes"
function convertOptions(data) {
  const options = {};
  if (data["Opção A"] !== undefined) options["Opção A"] = data["Opção A"];
  if (data["Opção B"] !== undefined) options["Opção B"] = data["Opção B"];
  if (data["Opção C"] !== undefined) options["Opção C"] = data["Opção C"];
  if (data["Opção D"] !== undefined) options["Opção D"] = data["Opção D"];

  delete data["Opção A"];
  delete data["Opção B"];
  delete data["Opção C"];
  delete data["Opção D"];

  if (Object.keys(options).length > 0) {
    data.opcoes = options;
  }

  return data;
}

// Função para verificar se o arquivo Excel corresponde ao modelo esperado
function verifyExcelFile(workbook) {
  const expectedSheetName = "Planilha1"; // Nome esperado da planilha

  // Verificar o nome da planilha
  if (!workbook.Sheets[expectedSheetName]) {
    throw new Error(
      "O arquivo Excel não corresponde ao modelo esperado: nome da planilha incorreto."
    );
  }
}

// Função principal para ler o arquivo Excel e convertê-lo para JSON
async function excelToJson(excelFilePath, jsonFilePath) {
  try {
    // Ler o arquivo Excel
    const workbook = XLSX.readFile(excelFilePath);

    // Verificar se o arquivo Excel corresponde ao modelo esperado
    verifyExcelFile(workbook);

    // Converter a planilha para JSON
    const sheetName = "Planilha1";
    const sheet = workbook.Sheets[sheetName];
    let jsonData = XLSX.utils.sheet_to_json(sheet, { defval: null });

    // Remover propriedades com valores null e converter opções
    jsonData = jsonData.map((data) =>
      convertOptions(removeNullProperties(data))
    );

    // Salvar o JSON formatado em um arquivo
    await fs.writeFile(jsonFilePath, JSON.stringify(jsonData, null, 2));
    console.log(`Arquivo JSON salvo em: ${jsonFilePath}`);
  } catch (error) {
    console.error(`Erro ao processar o arquivo: ${error.message}`);
  }
}

// Caminho do arquivo Excel e do arquivo JSON de saída
const excelFilePath = "./ModeloCreatorAuthor.xlsx";
const jsonFilePath = "./dadosSheet.json";

// Chamar a função para realizar a conversão
excelToJson(excelFilePath, jsonFilePath);

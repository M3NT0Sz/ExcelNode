const XLSX = require("xlsx");
const fs = require("fs");
const ExcelJS = require("exceljs");

// Função para remover propriedades com valores null ou undefined de um objeto
function removeNullProperties(obj) {
  return Object.fromEntries(Object.entries(obj).filter(([_, v]) => v != null));
}

// Função para construir o objeto rowData com base no tipo e na linha
function constructRowData(worksheet, row, tipo) {
  const baseData = {
    tipo: tipo,
    enunciado: row.getCell(2).value,
    resposta: row.getCell(3).value,
  };

  // Se o tipo tiver opções, adiciona-as ao objeto
  if (["Multipla Escolha", "Organizar", "Arrasta e Solta", "Associar", "Jogo da memória"].includes(tipo)) {
    baseData.opcoes = {
      opcaoA: row.getCell(4).value,
      opcaoB: row.getCell(5).value,
      opcaoC: row.getCell(6).value,
      opcaoD: row.getCell(7).value,
    };
    baseData.opcoes = removeNullProperties(baseData.opcoes);
  }

  return removeNullProperties(baseData);
}

// Função principal para ler o arquivo Excel e converter para JSON
async function excelToJson(excelFilePath, jsonFilePath) {
  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(excelFilePath);

    const worksheet = workbook.getWorksheet(1); // Obter a primeira planilha
    const jsonData = { licao: [] };

    // Iterar sobre as linhas da planilha, começando da segunda linha para evitar a primeira linha (cabeçalho)
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const row = worksheet.getRow(i);
      const tipo = row.getCell(1).value;

      // Verificar se a linha contém os cabeçalhos e pular essa linha
      if (tipo === "Tipo" || row.getCell(2).value === "Enunciado" || row.getCell(3).value === "Resposta") {
        continue;
      }

      if (tipo) {
        const rowData = constructRowData(worksheet, row, tipo);
        if (Object.keys(rowData).length > 0) {
          jsonData.licao.push(rowData);
        }
      }
    }

    // Escrever o JSON em um arquivo
    fs.writeFileSync(jsonFilePath, JSON.stringify(jsonData, null, 2));
    console.log(`Arquivo JSON salvo em ${jsonFilePath}`);
  } catch (err) {
    console.error("Erro ao processar o arquivo:", err);
  }
}

// Caminho do arquivo Excel e do arquivo JSON de saída
const excelFilePath = "./ModeloCreatorAuthor.xlsx";
const jsonFilePath = "./dadosSheet.json";

// Chamar a função para realizar a conversão
excelToJson(excelFilePath, jsonFilePath);

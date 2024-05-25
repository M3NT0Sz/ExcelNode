const XLSX = require("xlsx");
const fs = require("fs");

// Ler o arquivo Excel
const workbook = XLSX.readFile("./ModeloCreatorAuthor.xlsx");

// Obter a primeira planilha
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Converter a planilha para JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Obter cabeçalhos e linhas
const headers = jsonData[0];
const rows = jsonData.slice(1);

// Filtrar linhas onde o primeiro valor (tipo) não é nulo
const filteredJson = rows.filter((row) => row[0] !== null);

// Mapear as linhas para o formato desejado
const adjustedJson = filteredJson.map((row) => ({
  tipo: row[0],
  enunciado: row[1],
  resposta: row[3],
  opcaoA: row[4],
  opcaoB: row[5],
  opcaoC: row[6],
  opcaoD: row[7],
}));

// Remover objetos vazios ou com todas as propriedades indefinidas
const finalJson = adjustedJson.filter(obj => Object.values(obj).some(val => val !== undefined && val !== null));

// Salvar os dados em um arquivo JSON
fs.writeFileSync("dados.json", JSON.stringify(finalJson, null, 4));
console.log("Arquivo JSON gerado com sucesso!");

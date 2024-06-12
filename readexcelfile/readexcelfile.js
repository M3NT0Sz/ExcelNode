const XLSX = require('xlsx');
const fs = require('fs');

// Carregar o arquivo Excel
const workbook = XLSX.readFile('ModeloCreatorAuthor.xlsx');

// Obter o nome da primeira planilha
const sheetName = workbook.SheetNames[0];

// Obter a planilha pelo nome
const sheet = workbook.Sheets[sheetName];

// Converter a planilha para um objeto JSON
const jsonData = XLSX.utils.sheet_to_json(sheet);

// Remover a primeira linha (linhas 2 a 10)
jsonData.splice(0, 1);

// Renomear as chaves do objeto JSON
jsonData.forEach(row => {
  row['tipo'] = row['Lista de Atividades pelo Excel Creator Author'];
  delete row['Lista de Atividades pelo Excel Creator Author'];
  row['enunciado'] = row['__EMPTY'];
  delete row['__EMPTY'];
  row['resposta'] = row['__EMPTY_1'];
  delete row['__EMPTY_1'];
  row['opcaoA'] = row['__EMPTY_2'];
  delete row['__EMPTY_2'];
  row['opcaoB'] = row['__EMPTY_3'];
  delete row['__EMPTY_3'];
  row['opcaoC'] = row['__EMPTY_4'];
  delete row['__EMPTY_4'];
  row['opcaoD'] = row['__EMPTY_5'];
  delete row['__EMPTY_5'];
});

// Escrever os dados em um arquivo JSON
fs.writeFileSync('dadosRead.json', JSON.stringify(jsonData, null, 2));

console.log('Dados salvos em JSON com sucesso.');
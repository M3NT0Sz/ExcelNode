const XLSX = require("xlsx");
const fs = require("fs");
const ExcelJS = require("exceljs");

// Função para remover propriedades com valores null ou undefined de um objeto
function removeNullProperties(obj) {
  return Object.fromEntries(
    Object.entries(obj).filter(([_, v]) => v != null)
  );
}

// Ler o arquivo Excel com ExcelJS
const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile("./ModeloCreatorAuthor.xlsx")
  .then(() => {
    const worksheet = workbook.getWorksheet(1); // Obter a primeira planilha

    const jsonData = {
      "licao": []
    };

    // Iterar sobre as linhas da planilha, começando da segunda linha para evitar a primeira linha (cabeçalho)
    for (let i = 2; i <= worksheet.rowCount; i++) {
      let rowData = {};
      let tipo = worksheet.getRow(i).getCell(1).value;
      
      if (tipo === "Multipla Escolha") {
        rowData = {
          tipo: tipo,
          enunciado: worksheet.getRow(i).getCell(2).value,
          resposta: worksheet.getRow(i).getCell(3).value,
          opcoes: {
            opcaoA: worksheet.getRow(i).getCell(4).value,
            opcaoB: worksheet.getRow(i).getCell(5).value,
            opcaoC: worksheet.getRow(i).getCell(6).value,
            opcaoD: worksheet.getRow(i).getCell(7).value
          }
        };
        rowData.opcoes = removeNullProperties(rowData.opcoes);
      } else if (tipo === "Preenchimento") {
        rowData = {
          tipo: tipo,
          enunciado: worksheet.getRow(i).getCell(2).value,
          resposta: worksheet.getRow(i).getCell(3).value
        };
      } else if (tipo )

      rowData = removeNullProperties(rowData);

      // Adicionar o objeto de dados ao array de dados se não estiver vazio
      if (Object.keys(rowData).length > 0) {
        jsonData.licao.push(rowData);
      }
    }

    // Escrever o JSON em um arquivo
    fs.writeFileSync('dadosSheet.json', JSON.stringify(jsonData, null, 2));
    console.log('Arquivo JSON salvo em dadosSheet.json');
  })
  .catch(err => {
    console.error('Erro ao processar o arquivo:', err);
  });

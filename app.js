const ExcelJS = require("exceljs");
const fs = require("fs");

// Carregar o arquivo Excel
const workbook = new ExcelJS.Workbook();
workbook.xlsx
  .readFile("./ModeloCreatorAuthor.xlsx")
  .then(() => {
    // Assume que o arquivo tem apenas uma planilha, caso contrário, você precisará iterar sobre as planilhas
    const worksheet = workbook.getWorksheet(1);

    // Converter os dados da planilha para JSON
    const jsonData = [];
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (rowNumber !== 1) {
        const rowData = {};
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          rowData[`col${colNumber}`] = cell.value;
        });
        jsonData.push(rowData);
      }
    });

    // Filtrar as linhas com base no valor da coluna "tipo"
    const filteredJson = jsonData.filter((row) => {
      return row.col1 !== null; // Manter a linha se o valor da coluna "tipo" não for null
    });

    // Transformar o JSON filtrado em um formato ajustado
    const adjustedJson = filteredJson.map((row) => ({
      tipo: row.col1,
      enunciado: row.col2,
      imagemEnunciado: row.col3,
      resposta: row.col4,
      opcaoA: row.col5,
      opcaoB: row.col6,
      opcaoC: row.col7,
      opcaoD: row.col8,
      // Adicione mais campos conforme necessário
    }));

    // Salvar os dados em um arquivo JSON
    fs.writeFileSync("dados.json", JSON.stringify(adjustedJson, null, 4));
    console.log("Arquivo JSON gerado com sucesso!");
  })
  .catch((error) => {
    console.error("Erro ao ler o arquivo Excel:", error);
  });

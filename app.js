const ExcelJS = require("exceljs");
const fs = require("fs");

// Ler o arquivo Excel com exceljs
const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile("./ModeloCreatorAuthor.xlsx")
  .then(() => {
    const worksheet = workbook.getWorksheet(1); // Obter a primeira planilha

    const jsonData = {
      "licao": []
    };

    // Iterar sobre as linhas da planilha, começando da segunda linha para evitar a primeira linha (cabeçalho)
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const rowData = {};

      // Iterar sobre as células da linha
      worksheet.getRow(i).eachCell((cell, colIndex) => {
        switch (colIndex) {
          case 1:
            if (cell.value === "Multipla Escolha") {
              rowData.multiplaEscolha = { tipo: cell.value };
            } else if (cell.value === "Preenchimento") {
              rowData.preenchimento = { tipo: cell.value };
            }
            break;
          case 2:
            if (rowData.multiplaEscolha) {
              rowData.multiplaEscolha.enunciado = cell.value;
            }
            if (rowData.preenchimento) {
              rowData.preenchimento.enunciado = cell.value;
            }
            break;
          case 3:
            break;
          case 4:
            if (rowData.multiplaEscolha) {
              rowData.multiplaEscolha.resposta = cell.value;
            }
            if (rowData.preenchimento) {
              rowData.preenchimento.resposta = cell.value;
            }
            break;
          case 5:
            if (rowData.multiplaEscolha) {
              rowData.multiplaEscolha.opcoes = { opcaoA: cell.value };
            }
            break;
          case 6:
            if (rowData.multiplaEscolha) {
              rowData.multiplaEscolha.opcoes.opcaoB = cell.value;
            }
            break;
          case 7:
            if (rowData.multiplaEscolha) {
              rowData.multiplaEscolha.opcoes.opcaoC = cell.value;
            }
            break;
          case 8:
            if (rowData.multiplaEscolha) {
              rowData.multiplaEscolha.opcoes.opcaoD = cell.value;
            }
            break;
        }
      });

      // Adicionar o objeto de dados ao array de dados
      jsonData.licao.push(rowData);
    }

    // Salvar os dados em um arquivo JSON
    fs.writeFileSync("dados.json", JSON.stringify(jsonData, null, 4));
    console.log("Arquivo JSON gerado com sucesso!");
  })
  .catch((error) => {
    console.error("Erro ao ler o arquivo Excel:", error);
  });

const ExcelJS = require('exceljs');
const fs = require('fs');
const axios = require('axios');

async function excelToJson(filePath, sheetName) {
    // Cria uma nova instância do Workbook
    const workbook = new ExcelJS.Workbook();
    
    // Carrega o arquivo Excel
    await workbook.xlsx.readFile(filePath);

    // Seleciona a planilha
    const worksheet = workbook.getWorksheet(sheetName);

    // Extrai os dados
    const data = [];
    worksheet.eachRow(async (row, rowNumber) => {
        if (rowNumber > 1) { // Skip the header row
            const rowData = {
                Tipo: row.getCell(1).value,
                Enunciado: row.getCell(2).value,
                Resposta: row.getCell(3).value,
                OpcaoA: row.getCell(4).value,
                OpcaoB: row.getCell(5).value,
                OpcaoC: row.getCell(6).value,
                OpcaoD: row.getCell(7).value
            };
            // Check if there is a URL in any column for the image
            const imageUrl = row.getCell(8).value; // Assuming image URL is in column 8
            if (imageUrl) {
                const imageName = `image_${rowNumber}.png`;
                const imageResponse = await axios.get(imageUrl, { responseType: 'arraybuffer' });
                fs.writeFileSync(imageName, imageResponse.data);
                // Assign the image to the appropriate property (Tipo, Enunciado, OpcaoA, OpcaoB, OpcaoC, OpcaoD)
                const imageProperty = getImageProperty(rowData);
                rowData[imageProperty] = imageName; // Save image name instead of URL
            }
            data.push(rowData);
        }
    });

    // Converte para JSON
    const json = JSON.stringify(data, null, 2);

    // Salva o JSON em um arquivo
    fs.writeFileSync('dataExceljs.json', json, 'utf8');

    console.log('Arquivo JSON criado com sucesso!');
}

// Função auxiliar para determinar a propriedade a que a imagem deve ser atribuída
function getImageProperty(rowData) {
    const properties = ['Tipo', 'Enunciado', 'OpcaoA', 'OpcaoB', 'OpcaoC', 'OpcaoD'];
    for (const prop of properties) {
        if (!rowData[prop]) {
            return prop;
        }
    }
    return 'Imagem'; // Se todas as propriedades já estiverem ocupadas, atribuir à 'Imagem'
}

// Chama a função
excelToJson('../ModeloCreatorAuthor.xlsx', 'Planilha1').catch(console.error);

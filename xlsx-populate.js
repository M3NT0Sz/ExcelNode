const XlsxPopulate = require('xlsx-populate');
const fs = require('fs');

async function excelToJson(filePath, sheetName) {
    // Carrega o arquivo Excel
    const workbook = await XlsxPopulate.fromFileAsync(filePath);

    // Seleciona a planilha
    const worksheet = workbook.sheet(sheetName);

    // Extrai os dados
    const data = [];
    worksheet.usedRange().value().forEach((row, rowIndex) => {
        if (rowIndex > 0) { // Skip the header row
            const rowData = {
                Tipo: row[0],
                Enunciado: row[1],
                Resposta: row[2],
                OpcaoA: row[3],
                OpcaoB: row[4],
                OpcaoC: row[5],
                OpcaoD: row[6]
            };
            // Check for embedded image URL in any cell
            const imageUrl = getEmbeddedImageUrl(worksheet, rowIndex + 1); // Corrigido para indexação baseada em 1
            if (imageUrl) {
                rowData.Imagem = imageUrl; // Save image URL
            }
            data.push(rowData);
        }
    });

    // Converte para JSON
    const json = JSON.stringify(data, null, 2);

    // Salva o JSON em um arquivo
    fs.writeFileSync('dataPopulate.json', json, 'utf8');

    console.log('Arquivo JSON criado com sucesso!');
}

// Função para obter URL de imagem incorporada de uma linha
function getEmbeddedImageUrl(worksheet, rowIndex) {
    const cell = worksheet.cell(rowIndex, 1); // Corrigido para indexação baseada em 1
    if (cell._hyperlinks && cell._hyperlinks.length > 0) {
        const hyperlink = cell._hyperlinks[0];
        if (hyperlink.type === 'image') {
            return hyperlink.url;
        }
    }
    return null;
}

// Chama a função
excelToJson('ModeloCreatorAuthor.xlsx', 'Planilha1').catch(console.error);

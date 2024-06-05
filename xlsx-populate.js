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
            // Check for embedded image in any cell
            const embeddedImage = getEmbeddedImage(worksheet, rowIndex + 1); // Corrigido para indexação baseada em 1
            if (embeddedImage) {
                const imageName = `image_${rowIndex}.png`;
                fs.writeFileSync(imageName, embeddedImage);
                rowData.Imagem = imageName; // Save image name
            }
            data.push(rowData);
        }
    });

    // Converte para JSON
    const json = JSON.stringify(data, null, 2);

    // Salva o JSON em um arquivo
    fs.writeFileSync('data.json', json, 'utf8');

    console.log('Arquivo JSON criado com sucesso!');
}

// Função para obter imagem incorporada de uma linha
function getEmbeddedImage(worksheet, rowIndex) {
    const cell = worksheet.row(rowIndex).cell(1); // Corrigido para indexação baseada em 1
    if (cell.hasImage()) {
        return cell.image().toBuffer();
    }
    return null;
}

// Chama a função
excelToJson('Modelo', 'Planilha1').catch(console.error);

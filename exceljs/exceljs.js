const ExcelJS = require('exceljs');
const fs = require('fs');
const axios = require('axios');
const path = require('path');

async function excelToJson(filePath, sheetName) {
    // Cria uma nova instância do Workbook
    const workbook = new ExcelJS.Workbook();
    
    // Carrega o arquivo Excel
    await workbook.xlsx.readFile(filePath);

    // Seleciona a planilha
    const worksheet = workbook.getWorksheet(sheetName);

    // Extrai os dados e prepara para coletar imagens
    const data = [];
    const imageMap = {}; // Mapeia a linha da imagem para o nome do arquivo da imagem

    // Função para processar cada linha
    const processRow = async (row, rowNumber) => {
        if (rowNumber > 1) { // Pular a linha do cabeçalho
            const rowData = {
                Tipo: row.getCell(1).value,
                Enunciado: row.getCell(2).value,
                Resposta: row.getCell(3).value,
                OpcaoA: row.getCell(4).value,
                OpcaoB: row.getCell(5).value,
                OpcaoC: row.getCell(6).value,
                OpcaoD: row.getCell(7).value
            };

            // Verificar se há uma URL em alguma coluna para a imagem
            const imageUrl = row.getCell(8).value; // Supondo que o URL da imagem está na coluna 8
            if (imageUrl) {
                const imageName = `image_${rowNumber}.png`;
                try {
                    const imageResponse = await axios.get(imageUrl, { responseType: 'arraybuffer' });
                    fs.writeFileSync(imageName, imageResponse.data);
                    // Atribuir a imagem à propriedade apropriada (Tipo, Enunciado, OpcaoA, OpcaoB, OpcaoC, OpcaoD)
                    const imageProperty = getImageProperty(rowData);
                    rowData[imageProperty] = path.resolve(imageName); // Salva o caminho absoluto da imagem
                } catch (error) {
                    console.error(`Erro ao baixar a imagem ${imageUrl}:`, error.message);
                }
            }

            // Mapear imagens incorporadas
            const images = worksheet.getImages();
            for (const image of images) {
                if (image.range.tl.nativeRow === rowNumber) {
                    const img = workbook.model.media.find(m => m.index === image.imageId);
                    if (img) {
                        const imageFileName = `embedded_${rowNumber}_${image.range.tl.nativeCol}.${img.name}.${img.extension}`;
                        fs.writeFileSync(imageFileName, img.buffer);
                        // Atribuir o caminho da imagem ao campo apropriado
                        rowData[getImageProperty(rowData)] = path.resolve(imageFileName); // Atualiza o caminho
                    }
                }
            }

            data.push(rowData);
        }
    };

    // Processar cada linha da planilha
    for (let i = 1; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        await processRow(row, i);
    }

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

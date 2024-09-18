const ExcelJS = require('exceljs');
const fs = require('fs');
const axios = require('axios');
const path = require('path');

async function excelToJson(filePath, sheetName) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);

    const data = [];
    const imageMap = {};

    const processRow = async (row, rowNumber) => {
        if (rowNumber > 1) {
            const rowData = {
                Tipo: row.getCell(1).value,
                Enunciado: row.getCell(2).value,
                Resposta: row.getCell(3).value,
                OpcaoA: row.getCell(4).value,
                OpcaoB: row.getCell(5).value,
                OpcaoC: row.getCell(6).value,
                OpcaoD: row.getCell(7).value,
                Imagens: {}
            };

            const imageUrl = row.getCell(8).value;
            if (imageUrl) {
                const imageName = `image_${rowNumber}.png`;
                try {
                    const imageResponse = await axios.get(imageUrl, { responseType: 'arraybuffer' });
                    fs.writeFileSync(imageName, imageResponse.data);
                    rowData.Imagens['URL'] = path.resolve(imageName);
                } catch (error) {
                    console.error(`Erro ao baixar a imagem ${imageUrl}:`, error.message);
                }
            }

            const images = worksheet.getImages();
            for (const image of images) {
                if (image.range.tl.nativeRow === rowNumber - 1) {
                    const img = workbook.model.media.find(m => m.index === image.imageId);
                    if (img) {
                        const imageFileName = `embedded_${rowNumber}_${image.range.tl.nativeCol}.${img.name}.${img.extension}`;
                        fs.writeFileSync(imageFileName, img.buffer);
                        rowData.Imagens[`Coluna${image.range.tl.nativeCol + 1}`] = path.resolve(imageFileName);
                    }
                }
            }

            data.push(rowData);
        }
    };

    for (let i = 1; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        await processRow(row, i);
    }

    const json = JSON.stringify(data, null, 2);
    fs.writeFileSync('dataExceljs.json', json, 'utf8');
    console.log('Arquivo JSON criado com sucesso!');
}

excelToJson('../ModeloCreatorAuthor.xlsx', 'Planilha1').catch(console.error);

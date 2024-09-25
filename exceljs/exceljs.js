const ExcelJS = require("exceljs");
const fs = require("fs");
const axios = require("axios");

async function excelToJson(filePath, sheetName) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);

    const data = [];

    const processRow = async (row, rowNumber) => {
        if (rowNumber > 1) {
            const rowData = {
                id: rowNumber.toString(),
                tipo: row.getCell(1).value || "",
                contexto: [
                    {
                        tipo: "texto",
                        dados: {
                            html: row.getCell(2).value || "",
                            audio: "",
                            centralizado: "false",
                        },
                    },
                ],
                alternativas: [
                    {
                        id: "A",
                        img: "",
                        descricao: row.getCell(4).value || "",
                        distrator: "",
                        "exibir-descricao": "false",
                        resposta:
                            row.getCell(3).value === row.getCell(4).value
                                ? "correta"
                                : "incorreta",
                    },
                    {
                        id: "B",
                        img: "",
                        descricao: row.getCell(5).value || "",
                        distrator: "",
                        "exibir-descricao": "false",
                        resposta:
                            row.getCell(3).value === row.getCell(5).value
                                ? "correta"
                                : "incorreta",
                    },
                    {
                        id: "C",
                        img: "",
                        descricao: row.getCell(6).value || "",
                        distrator: "",
                        "exibir-descricao": "false",
                        resposta:
                            row.getCell(3).value === row.getCell(6).value
                                ? "correta"
                                : "incorreta",
                    },
                    {
                        id: "D",
                        img: "",
                        descricao: row.getCell(7).value || "",
                        distrator: "",
                        "exibir-descricao": "false",
                        resposta:
                            row.getCell(3).value === row.getCell(7).value
                                ? "correta"
                                : "incorreta",
                    },
                ],
                feedback: {
                    tipo: "opcao3",
                    dados: {},
                },
                "is-professor": "false",
            };

            const imageUrl = row.getCell(8).value;
            if (imageUrl) {
                const imageName = `image_${rowNumber}.png`;
                try {
                    const imageResponse = await axios.get(imageUrl, {
                        responseType: "arraybuffer",
                    });
                    fs.writeFileSync(imageName, imageResponse.data);
                    rowData.contexto[0].dados.img = imageName;
                } catch (error) {
                    console.error(`Erro ao baixar a imagem ${imageUrl}:`, error.message);
                }
            }

            const images = worksheet.getImages();
            for (const image of images) {
                if (image.range.tl.nativeRow === rowNumber - 1) {
                    const img = workbook.model.media.find(
                        (m) => m.index === image.imageId
                    );
                    if (img) {
                        const colIndex = image.range.tl.nativeCol - 3;
                        if (colIndex >= 0 && colIndex < rowData.alternativas.length) {
                            const imageFileName = `embedded_${rowNumber}_${image.range.tl.nativeCol}.${img.name}.${img.extension}`;
                            fs.writeFileSync(imageFileName, img.buffer);
                            rowData.alternativas[colIndex].img = imageFileName;
                        }
                    }
                }
            }

            // Check if rowData matches the specified structure
            const isEmptyRowData =
                rowData.tipo === "" &&
                rowData.contexto[0].dados.html === "" &&
                rowData.alternativas.every((alt) => alt.descricao === "");

            if (!isEmptyRowData) {
                data.push(rowData);
            }
        }
    };

    for (let i = 1; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        await processRow(row, i);
    }

    if (data.length > 0) {
        const json = JSON.stringify(data, null, 2);
        fs.writeFileSync("dataExceljs.json", json, "utf8");
        console.log("Arquivo JSON criado com sucesso!");

        data.forEach((rowData, index) => {
            const json = JSON.stringify(rowData, null, 2);
            fs.writeFileSync(`dataExceljs_${index + 1}.json`, json, "utf8");

            // Verificar qual alternativa está correta
            const correctAlternative = rowData.alternativas.find(
                (alt) => alt.resposta === "correta"
            );
            if (correctAlternative) {
                console.log(
                    `A alternativa correta para a questão ${rowData.id} é: ${correctAlternative.id}`
                );
            } else {
                console.log(`Nenhuma alternativa correta encontrada para a questão ${rowData.id}`);
            }
        });
    } else {
        console.log("Nenhum dado encontrado para criar arquivos JSON.");
    }
}

excelToJson("../ModeloCreatorAuthor.xlsx", "Planilha1").catch(console.error);

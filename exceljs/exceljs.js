const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

async function excelToJson(filePath, sheetName) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet(sheetName);

    const data = [];

    const processRow = async (row, rowNumber) => {
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
                        img: ""
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

        // Get the images in the worksheet
        const images = worksheet.getImages();
        for (const image of images) {
            if (image.range.tl.nativeRow === rowNumber - 1) {
                const img = workbook.model.media.find(
                    (m) => m.index === image.imageId
                );
                if (img) {
                    const questionDir = path.join(__dirname, rowData.id);
                    if (!fs.existsSync(questionDir)) {
                        fs.mkdirSync(questionDir);
                    }
                    const imageFileName = `embedded_${rowNumber}_${image.range.tl.nativeCol}.${img.extension}`;
                    const imageFilePath = path.join(questionDir, imageFileName);
                    fs.writeFileSync(imageFilePath, img.buffer);

                    // Check if the image is associated with an alternative
                    const alternativeIndex = image.range.tl.nativeCol - 3;
                    if (alternativeIndex >= 0 && alternativeIndex < rowData.alternativas.length) {
                        rowData.alternativas[alternativeIndex].img = imageFileName;
                    } else if (image.range.tl.nativeCol === 2) { // Assuming column 2 is for context image
                        rowData.contexto[0].dados.img = imageFileName;
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
    };

    for (let i = 1; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        await processRow(row, i);
    }

    if (data.length > 0) {
        data.forEach((rowData, index) => {
            const questionDir = path.join(__dirname, rowData.id);
            if (!fs.existsSync(questionDir)) {
                fs.mkdirSync(questionDir);
            }
            const json = JSON.stringify(rowData, null, 2);
            fs.writeFileSync(path.join(questionDir, `dataExceljs_${index + 1}.json`), json, "utf8");

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

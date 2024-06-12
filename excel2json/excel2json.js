const XLSX = require('xlsx');
const fs = require('fs');

// Caminho do arquivo Excel
const excelFilePath = 'ModeloCreatorAuthor.xlsx';

// Função para ler o arquivo Excel e organizar os dados
function readAndOrganizeExcel(filePath) {
    // Carrega o arquivo Excel
    const workbook = XLSX.readFile(filePath);

    // Obtém a primeira planilha
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Converte a planilha para JSON
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Organiza os dados conforme o formato desejado
    const organizedData = {
        licao: []
    };

    // Itera sobre os dados extraídos do Excel
    for (let i = 1; i < jsonData.length; i++) {
        const row = jsonData[i];

        // Cria um objeto para cada linha de dados
        const item = {
            tipo: row[0],
            enunciado: row[1],
            resposta: row[2]
        };

        // Verifica se há opções múltiplas e as organiza
        if (item.tipo === 'Multipla Escolha') {
            item.opcoes = {
                opcaoA: row[3],
                opcaoB: row[4],
                opcaoC: row[5],
                opcaoD: row[6]
            };
        }

        // Adiciona o item à lista de lições
        organizedData.licao.push(item);
    }

    return organizedData;
}

// Verifica se o arquivo existe
if (fs.existsSync(excelFilePath)) {
    // Chama a função para ler e organizar o arquivo Excel
    const organizedJsonData = readAndOrganizeExcel(excelFilePath);

    // Caminho do arquivo JSON onde os dados organizados serão salvos
    const jsonFilePath = 'dadosexcel2json.json';

    // Converte os dados organizados para uma string JSON formatada
    const jsonString = JSON.stringify(organizedJsonData, null, 2);

    // Escreve os dados no arquivo JSON
    fs.writeFile(jsonFilePath, jsonString, (err) => {
        if (err) {
            console.error('Erro ao salvar o arquivo JSON organizado:', err);
        } else {
            console.log(`Dados organizados salvos com sucesso em ${jsonFilePath}`);
        }
    });
} else {
    console.error(`Arquivo não encontrado: ${excelFilePath}`);
}

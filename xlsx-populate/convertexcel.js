const convertExcelToJson = require('convert-excel-to-json');
const fs = require('fs');

// Caminho do arquivo Excel
const filePath = 'ModeloCreatorAuthor.xlsx';

// Verifica se o arquivo existe
if (fs.existsSync(filePath)) {
    const result = convertExcelToJson({
        sourceFile: filePath,
        header: {
            rows: 1
        },
        // Mapeamento de colunas específicas
        columnToKey: {
            A: 'Lista de Atividades pelo Excel Creator Author', // Suponha que a coluna A contém esses valores
            B: 'Tipo', // Suponha que a coluna B contém os valores "Tipo"
            C: 'resposta',
            D: 'opção A',
            E: 'opção B',
            F: 'opção C',
            G: 'opção D'// Adicione mais mapeamentos conforme necessário
        }
    });

    // Caminho do arquivo JSON onde os dados serão salvos
    const jsonFilePath = 'dadosexcel.json';

    // Converte os dados para uma string JSON formatada
    const jsonString = JSON.stringify(result, null, 2);

    // Escreve os dados no arquivo JSON
    fs.writeFile(jsonFilePath, jsonString, (err) => {
        if (err) {
            console.error('Erro ao salvar o arquivo JSON:', err);
        } else {
            console.log(`Dados salvos com sucesso em ${jsonFilePath}`);
        }
    });
} else {
    console.error(`Arquivo não encontrado: ${filePath}`);
}
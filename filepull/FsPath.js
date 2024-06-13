const fs = require('fs');
const path = require('path');

// Nome do arquivo que você deseja obter o caminho
let nomeDoArquivo = 'linkGitHub.txt';

// Obter o caminho absoluto do diretório atual
const diretorioAtual = process.cwd();

//Obter o caminho absoluto do arquivo
const caminhoDoArquivo = path.join(diretorioAtual, nomeDoArquivo);

console.log("Caminho encontrado: ", caminhoDoArquivo);
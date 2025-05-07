const fs = require('fs');
const path = require('path');

function contarPastas(diretorio) {
  try {
    const itens = fs.readdirSync(diretorio, { withFileTypes: true });
    const pastas = itens.filter(item => item.isDirectory());
    console.log(`Número de pastas em "${diretorio}": ${pastas.length}`);
    return pastas.length;
  } catch (erro) {
    console.error(`Erro ao ler o diretório: ${erro.message}`);
    return 0;
  }
}

// Exemplo de uso
const caminho_pastas = 'H:/Meu Drive/Carpi/Cadastros/Cadastros completos'; // 35
contarPastas(caminho_pastas);

function contarArquivos(diretorio) {
  try {
    const itens = fs.readdirSync(diretorio, { withFileTypes: true });
    const arquivos = itens.filter(item => item.isFile());
    console.log(`Número de arquivos em "${diretorio}": ${arquivos.length}`);
    return arquivos.length;
  } catch (erro) {
    console.error(`Erro ao ler o diretório: ${erro.message}`);
    return 0;
  }
}

// Exemplo de uso
const caminho_arquivos = 'D:/G3Core/SistemaCarpi/Carpi_scripts/dados-funcionarios';
contarArquivos(caminho_arquivos);
const fs = require('fs');
const path = require('path');

const pastaBase = 'H:/Meu Drive/Carpi/Cadastros/5. Cadastros Antigos de Funcionários';

function contarPastas(diretorio) {
  try {
    const itens = fs.readdirSync(diretorio, { withFileTypes: true });
    const pastas = itens.filter(item => item.isDirectory());
    return pastas.length;
  } catch (error) {
    console.error('Erro ao acessar o diretório:', error.message);
    return 0;
  }
}

const quantidadePastas = contarPastas(pastaBase);
console.log(`Quantidade de pastas em "${pastaBase}": ${quantidadePastas}`);
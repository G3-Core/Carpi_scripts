const fs = require('fs');
const path = require('path');

// Configuração de caminhos
const pastaBase = 'H:/Meu Drive/Carpi/Cadastros/5. Cadastros Antigos de Funcionários';
const pastaDestino = 'H:/Meu Drive/Carpi/Cadastros/Cadastros duplicados';

/**
 * Normaliza o nome removendo acentos, preposições, espaços extras e singularizando
 * @param {string} nome - Nome original da pasta
 * @return {string} Nome normalizado
 */
function normalizarNome(nome) {
  return nome
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "") // Remove acentos
    .replace(/\b(de|do|da|dos|das)\b/g, '') // Remove preposições/artigos
    .replace(/\s+/g, ' ') // Remove espaços extras
    .trim()
    .replace(/s\b/g, ''); // Reduz final plural
}

/**
 * Move uma pasta de origem para destino
 * @param {string} origem - Caminho completo da pasta de origem
 * @param {string} destino - Caminho completo da pasta de destino
 */
function moverPasta(origem, destino) {
  try {
    // Verifica se o destino já existe
    if (fs.existsSync(destino)) {
      // Cria um nome único adicionando timestamp
      const timestamp = Date.now();
      const dirName = path.basename(destino);
      const parentDir = path.dirname(destino);
      const novoPastaDestino = path.join(parentDir, `${dirName}_${timestamp}`);
      fs.renameSync(origem, novoPastaDestino);
      console.log(`Movido: ${origem} -> ${novoPastaDestino} (nome alterado para evitar conflito)`);
    } else {
      fs.renameSync(origem, destino);
      console.log(`Movido: ${origem} -> ${destino}`);
    }
  } catch (err) {
    console.error(`Erro ao mover "${origem}":`, err.message);
  }
}

/**
 * Processa todas as pastas, identifica duplicadas e move TODAS para o destino
 * @param {string} diretorio - Diretório base para buscar pastas
 * @param {string} destinoDuplicadas - Diretório onde serão armazenadas as duplicadas
 */
function identificarEMoverDuplicadas(diretorio, destinoDuplicadas) {
  try {
    // Garante que a pasta de destino exista
    if (!fs.existsSync(destinoDuplicadas)) {
      fs.mkdirSync(destinoDuplicadas, { recursive: true });
    }

    // Obtém todas as pastas no diretório
    const itens = fs.readdirSync(diretorio, { withFileTypes: true });
    const pastas = itens.filter(item => item.isDirectory());
    
    // Mapa para rastrear pastas com nomes normalizados semelhantes
    const mapaNomes = new Map();
    
    // Preenche o mapa com todas as pastas
    pastas.forEach(pasta => {
      const nomeOriginal = pasta.name;
      const nomeNormalizado = normalizarNome(nomeOriginal);
      
      if (!mapaNomes.has(nomeNormalizado)) {
        mapaNomes.set(nomeNormalizado, []);
      }
      mapaNomes.get(nomeNormalizado).push(nomeOriginal);
    });
    
    // Contadores para estatísticas
    let gruposDuplicados = 0;
    let pastasDuplicadasMovidas = 0;
    
    // Processa cada grupo de nomes normalizados
    for (const [nomeNormalizado, listaNomes] of mapaNomes.entries()) {
      // Se tiver mais de uma pasta com o mesmo nome normalizado
      if (listaNomes.length > 1) {
        gruposDuplicados++;
        console.log(`\nGrupo de duplicadas encontrado (${nomeNormalizado}):\n- ${listaNomes.join('\n- ')}`);
        
        // Move TODAS as pastas duplicadas
        listaNomes.forEach(nomeOriginal => {
          const caminhoOrigem = path.join(diretorio, nomeOriginal);
          const caminhoDestino = path.join(destinoDuplicadas, nomeOriginal);
          moverPasta(caminhoOrigem, caminhoDestino);
          pastasDuplicadasMovidas++;
        });
      }
    }
    
    console.log(`\n======= RESUMO =======`);
    console.log(`Grupos de duplicadas encontrados: ${gruposDuplicados}`);
    console.log(`Total de pastas movidas: ${pastasDuplicadasMovidas}`);
    console.log(`Todas as pastas foram movidas para: ${destinoDuplicadas}`);
    
  } catch (error) {
    console.error('Erro ao processar diretório:', error.message);
  }
}

// Executa o processamento
identificarEMoverDuplicadas(pastaBase, pastaDestino);
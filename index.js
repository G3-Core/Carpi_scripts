const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Caminho para a pasta onde estão os arquivos
const pastaBase = 'H:/Meu Drive/Carpi/Cadastros/5. Cadastros Antigos de Funcionários';

// Função para identificar o arquivo correto com "( I )" no nome
function obterArquivoFichaCompleta(arquivos) {
  for (const arquivo of arquivos) {
    const nomeArquivo = path.basename(arquivo);
    if (nomeArquivo.includes('( I )')) {
      return arquivo;
    }
  }
  // Se não encontrar "( I )", retorna o primeiro arquivo .xlsx encontrado
  return arquivos.length > 0 ? arquivos[0] : null;
}

// Função para extrair telefones de um texto
function extrairTelefones(texto) {
  if (!texto || typeof texto !== 'string') return [];
  
  const regexTelefone = /\(?(\d{2})\)?[\s-]?(\d{4,5})[\s-]?(\d{4})/g;
  const telefonesEncontrados = [];
  let match;
  
  while ((match = regexTelefone.exec(texto)) !== null) {
    const telefoneCompleto = match[0];
    const indiceNome = texto.indexOf('Nome:', match.index + telefoneCompleto.length);
    
    let nome = null;
    if (indiceNome !== -1) {
      const fimNome = texto.indexOf('\n', indiceNome);
      const finalNome = fimNome !== -1 ? fimNome : texto.length;
      nome = texto.substring(indiceNome + 5, finalNome).trim();
    }
    
    telefonesEncontrados.push({
      numero: telefoneCompleto,
      nome: nome
    });
  }
  
  return telefonesEncontrados;
}

// Função para calcular idade a partir da data de nascimento
function calcularIdade(dataTexto) {
  const match = dataTexto.match(/\d{2}\/\d{2}\/\d{4}/);
  if (!match) return dataTexto;
  
  const [dia, mes, ano] = match[0].split("/").map(Number);
  const nascimento = new Date(ano, mes - 1, dia);
  const hoje = new Date();
  
  let idade = hoje.getFullYear() - nascimento.getFullYear();
  const m = hoje.getMonth() - nascimento.getMonth();
  if (m < 0 || (m === 0 && hoje.getDate() < nascimento.getDate())) idade--;
  
  return `${idade} anos (${match[0]})`;
}

// Função para extrair telefones de recado
function extrairTelefonesRecado(dados) {
  const telefonesRecado = [];
  
  // Se já temos telefones processados
  if (dados.telefones_recado && Array.isArray(dados.telefones_recado)) {
    return dados.telefones_recado;
  }
  
  // Verificar se temos um telefone_para_recado e um nome associado
  if (dados.telefone_para_recado && dados.nome && 
      !dados.nome.includes('_') && !dados.nome_completo.includes(dados.nome)) {
    telefonesRecado.push({
      numero: dados.telefone_para_recado,
      nome: dados.nome
    });
  } else if (dados.telefone_para_recado) {
    // Extrair telefones do texto
    const telefonesExtraidos = extrairTelefones(dados.telefone_para_recado);
    if (telefonesExtraidos.length > 0) {
      telefonesRecado.push(...telefonesExtraidos);
    } else {
      // Se não encontrou nomes automáticos, adicione apenas o número
      telefonesRecado.push({
        numero: dados.telefone_para_recado,
        nome: null
      });
    }
  }
  
  // Processar telefones adicionais (telefone2, telefone3, etc.)
  for (let i = 2; i <= 10; i++) {
    const campoTelefone = `telefone${i}`;
    if (dados[campoTelefone]) {
      const telefonesExtraidos = extrairTelefones(dados[campoTelefone]);
      if (telefonesExtraidos.length > 0) {
        telefonesRecado.push(...telefonesExtraidos);
      } else {
        telefonesRecado.push({
          numero: dados[campoTelefone],
          nome: null
        });
      }
    }
  }
  
  return telefonesRecado;
}

// Função para agrupar os dados em categorias lógicas
function agruparDados(dados) {
  // Extrair telefones de recado
  const telefonesRecado = extrairTelefonesRecado(dados);
  
  // Determinar se tem telefone
  const temTelefone = Boolean(dados.telefone_pessoal || telefonesRecado.length > 0);
  
  return {
    nome_completo: dados.nome_completo,
    idade: dados.idade,
    pessoal: {
      fumante: dados.fumante,
      estado_civil: dados.estado_civil,
      naturalidade: dados.naturalidade,
      filhos: dados.filhos,
      saude: dados.sade || dados.saúde,
      religiao: dados.religio || dados.religião,
      conjuge: dados.cnjuge || dados.cônjuge
    },
    documentos: {
      rg: dados.rg,
      data_expedicao: dados.data_expedio || dados.data_expedição,
      cpf: dados.cpf,
      pis: dados.pis,
      carteira_num: dados.carteira_n || dados.carteira_nº,
      serie: dados.srie || dados.série
    },
    endereco: {
      rua: dados.endereo || dados.endereço,
      numero: dados.numero || dados["n°"],
      bairro: dados.bairro,
      cidade: dados.cidade,
      tempo_de_residencia: dados.tempo_de_residncia || dados.tempo_de_residência
    },
    contato: {
      tem_telefone: temTelefone,
      telefone_pessoal: dados.telefone_pessoal,
      telefones_recado: telefonesRecado
    },
    educacao: {
      escolaridade: dados.escolaridade
    },
    profissional: {
      profissao: dados.profisso || dados.profissão,
      cargo_desejado: dados.cargo_desejado,
      habilitacao: dados.habilitao || dados.habilitação,
      categoria: dados.categoria,
      veiculo_proprio: dados.veculo_prprio || dados.veículo_próprio,
      pretensao_salarial: dados.pretenso_salarial || dados.pretensão_salarial,
      ultimo_salario: dados.ltimo_salrio || dados.último_salário,
      disponibilidade_para_dormir: dados.disponibilidade_para_dormir,
      disponivel_aos_sabados: dados.disponvel_aos_sbados || dados.disponível_aos_sábados
    },
    habilidades: {
      lava_e_passa_todo_tipo_de_roupa: dados.lava_e_passa_todo_tipo_de_roupa,
      lava_e_passa_o_basico: dados.lava_e_passa_o_bsico || dados.lava_e_passa_o_básico
      // Adicionar outras habilidades específicas aqui
    },
    obs: dados.obs,
    referencias_profissionais: dados.referencias_profissionais
  };
}

// Função para processar um arquivo Excel e extrair os dados
function processarArquivoExcel(arquivo) {
  console.log(`Processando arquivo: ${arquivo}`);
  
  try {
    const workbook = xlsx.readFile(arquivo);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    
    let dados = {};
    let referencias = [];
    let referenciaAtual = null;
    let modoReferencia = false;
    let ultimaChave = null;
    
    // Normalizar chaves para evitar problemas com acentos e espaços
    function normalizarChave(chave) {
      return chave
        .toLowerCase()
        .replace(/[\s\.]+/g, '_')
        .replace(/[^\w_]/g, '');
    }
    
    // Verificar se um texto inicia uma nova referência profissional
    function ehNovaReferencia(texto) {
      return /^\d{1,2}°?\s*[-_]\s*nome/i.test(texto);
    }
    
    // Extrair índice da referência
    function extrairIndiceReferencia(texto) {
      const match = texto.match(/^(\d{1,2})°?\s*[-_]/i);
      return match ? match[1] : null;
    }
    
    // Processar cada linha da planilha
    for (const row of rows) {
      for (let j = 0; j < row.length; j++) {
        const cell = row[j];
        if (!cell || typeof cell !== 'string') continue;
        
        const texto = cell.trim();
        if (!texto) continue;
        
        // Verificar se entramos na seção de referências
        if (texto.toUpperCase().includes('REFERÊNCIA PROFISSIONAL')) {
          modoReferencia = true;
          continue;
        }
        
        // Processar células que contêm ":"
        if (texto.includes(':')) {
          const [rawKey, ...rawValueParts] = texto.split(':');
          const chaveOriginal = rawKey.trim();
          let chaveFormatada = normalizarChave(chaveOriginal);
          let valor = rawValueParts.join(':').trim();
          
          // Tratamento especial para algumas chaves
          if (chaveFormatada === 'n') chaveFormatada = 'numero';
          if (chaveFormatada === 'endereco') chaveFormatada = 'endereco';
          
          // Tratamento especial para idade
          if (!modoReferencia && chaveFormatada === 'idade' && valor.includes('/')) {
            valor = calcularIdade(valor);
          }
          
          ultimaChave = chaveFormatada;
          
          if (modoReferencia) {
            if (ehNovaReferencia(chaveOriginal)) {
              if (referenciaAtual) referencias.push(referenciaAtual);
              referenciaAtual = {};
              const indice = extrairIndiceReferencia(chaveOriginal);
              if (indice) referenciaAtual["indice"] = indice;
              referenciaAtual["nome"] = valor;
            } else {
              if (!referenciaAtual) referenciaAtual = {};
              referenciaAtual[chaveFormatada] = valor;
            }
          } else {
            dados[chaveFormatada] = valor;
          }
        } 
        // Para linhas que continuam um campo anterior
        else if (ultimaChave) {
          if (modoReferencia) {
            if (!referenciaAtual) referenciaAtual = {};
            referenciaAtual[ultimaChave] = (referenciaAtual[ultimaChave] || '') + ' ' + texto;
          } else {
            dados[ultimaChave] = (dados[ultimaChave] || '') + ' ' + texto;
          }
        }
      }
    }
    
    // Adicionar a última referência, se existir
    if (referenciaAtual) referencias.push(referenciaAtual);
    if (referencias.length > 0) dados["referencias_profissionais"] = referencias;
    
    // Criar saída no formato do exemplo (dados não agrupados)
    const dadosSaida = { ...dados };
    
    // Adicionar as referências profissionais
    if (referencias.length > 0) {
      dadosSaida.referencias_profissionais = referencias;
    }
    
    // Agrupar os dados em categorias lógicas
    const dadosAgrupados = agruparDados(dadosSaida);
    
    // Criar pasta de saída se não existir
    const pastaSaida = path.join(__dirname, 'dados-funcionarios-antigos');
    if (!fs.existsSync(pastaSaida)) {
      fs.mkdirSync(pastaSaida, { recursive: true });
    }
    
    // Gerar arquivo JSON com o nome do funcionário
    const nomeFuncionario = dados.nome_completo || path.basename(arquivo, '.xlsx');
    const nomeFormatado = nomeFuncionario.replace(/[^a-zA-Z0-9]/g, '-');
    
    // Salvar versão original (não agrupada) - nao vejo motivo para isso
    // const caminhoArquivoJsonOriginal = path.join(pastaSaida, `dados-brutos-${nomeFormatado}.json`);
    // fs.writeFileSync(caminhoArquivoJsonOriginal, JSON.stringify(dadosSaida, null, 2), 'utf8');
    
    // Salvar versão agrupada
    const caminhoArquivoJsonAgrupado = path.join(pastaSaida, `dados-${nomeFormatado}.json`);
    fs.writeFileSync(caminhoArquivoJsonAgrupado, JSON.stringify(dadosAgrupados, null, 2), 'utf8');
    
    console.log(`Arquivo JSON gerado: Dados agrupados: ${caminhoArquivoJsonAgrupado}`);
    
    return dadosAgrupados;
  } catch (error) {
    console.error(`Erro ao processar o arquivo ${arquivo}:`, error);
    return null;
  }
}

// Função para agrupar os arquivos por pasta e processar
function processarArquivos() {
  // Cria a pasta de saída se não existir
  const pastaSaida = path.join(__dirname, 'dados-funcionarios-antigos');
  if (!fs.existsSync(pastaSaida)) {
    fs.mkdirSync(pastaSaida, { recursive: true });
  }
  
  // Obtém todas as pastas no diretório base
  const pastas = fs.readdirSync(pastaBase)
    .map(item => path.join(pastaBase, item))
    .filter(item => fs.statSync(item).isDirectory());
  
  // Limita a quantidade de pastas processadas
  const pastasLimitadas = pastas.slice(0, 20); // limite de pastas
  
  console.log(`Total de pastas encontradas: ${pastas.length}`);
  console.log(`Processando ${pastasLimitadas.length} pastas`);
  
  for (const pasta of pastasLimitadas) {
    console.log(`\nProcessando pasta: ${pasta}`);
    
    // Obtém todos os arquivos .xlsx na pasta atual (sem subpastas)
    const arquivos = fs.readdirSync(pasta)
      .filter(item => item.endsWith('.xlsx'))
      .map(item => path.join(pasta, item));
    
    if (arquivos.length === 0) {
      console.log(`Nenhum arquivo Excel encontrado na pasta ${pasta}`);
      continue;
    }
    
    console.log(`Encontrados ${arquivos.length} arquivos Excel`);
    
    // Identifica o arquivo com "( I )" no nome
    const arquivoFichaCompleta = obterArquivoFichaCompleta(arquivos);
    
    if (arquivoFichaCompleta) {
      processarArquivoExcel(arquivoFichaCompleta);
    } else {
      console.log(`Nenhum arquivo de ficha completa encontrado na pasta ${pasta}`);
    }
  }
  
  console.log('\nProcessamento concluído!');
}

// Executa o processamento principal
processarArquivos();
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

// Caminho para a pasta onde estão os arquivos
const pastaBase = 'H:/Meu Drive/Carpi/Cadastros/5. Cadastros Antigos de Funcionários';

// Função para obter os arquivos Excel na pasta e subpastas, mas limitar para as duas primeiras pastas
function obterArquivosNaPasta(pasta, limitePastas = 2) {
  let arquivos = [];
  let pastasEncontradas = 0;
  
  // Ler o conteúdo da pasta
  const itens = fs.readdirSync(pasta);
  
  // Filtra apenas arquivos e subpastas
  itens.forEach(item => {
    const caminhoItem = path.join(pasta, item);
    const stats = fs.statSync(caminhoItem);
    
    if (stats.isDirectory() && pastasEncontradas < limitePastas) {
      // Incrementa o contador de pastas encontradas
      pastasEncontradas++;
      // Recursivamente obter arquivos das subpastas
      arquivos = arquivos.concat(obterArquivosNaPasta(caminhoItem, limitePastas));
    } else if (item.endsWith('.xlsx')) {
      arquivos.push(caminhoItem);
    }
  });

  return arquivos;
}

// Função para encontrar o arquivo com a versão mais recente
function obterArquivoMaisRecente(arquivos) {
  let arquivoMaisRecente = null;
  let maiorVersao = 0;

  arquivos.forEach(arquivo => {
    const nomeArquivo = path.basename(arquivo);
    
    // Procurar a versão romana no nome do arquivo
    const match = nomeArquivo.match(/\( (I|II|III|IV|V|VI|VII|VIII|IX|X) \)/);
    if (match) {
      const versao = match[1];
      const versaoNumerica = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X'].indexOf(versao) + 1;
      
      if (versaoNumerica > maiorVersao) {
        maiorVersao = versaoNumerica;
        arquivoMaisRecente = arquivo;
      }
    }
  });

  return arquivoMaisRecente;
}

// Função para processar um arquivo Excel e gerar JSON
function processarArquivoExcel(arquivo) {
  const workbook = xlsx.readFile(arquivo);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  let dados = {};
  let referencias = [];
  let referenciaAtual = null;
  let modoReferencia = false;
  let ultimaChave = null;

  const ehNovaReferencia = (texto) => /^\d{1,2}°?\s*[-_]\s*nome/i.test(texto);

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

  function extrairIndiceReferencia(texto) {
    const match = texto.match(/^(\d{1,2})°?\s*[-_]/i);
    return match ? match[1] : null;
  }

  // Processa a planilha
  for (const row of rows) {
    for (const cell of row) {
      if (!cell || typeof cell !== "string") continue;
      const texto = cell.trim();
      if (!texto) continue;

      if (texto.toUpperCase().includes("REFERÊNCIA PROFISSIONAL")) {
        modoReferencia = true;
        continue;
      }

      if (texto.includes(":")) {
        const [rawKey, ...rawValue] = texto.split(":");
        const chaveOriginal = rawKey.trim();
        const chaveFormatada = chaveOriginal.toLowerCase().replace(/[\s\.]+/g, "_");
        const valor = rawValue.join(":").trim();

        if (!modoReferencia && chaveFormatada === "idade") {
          dados["idade"] = calcularIdade(valor);
          ultimaChave = "idade";
          continue;
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
      } else if (ultimaChave) {
        if (modoReferencia) {
          if (!referenciaAtual) referenciaAtual = {};
          referenciaAtual[ultimaChave] += " " + texto;
        } else {
          dados[ultimaChave] += " " + texto;
        }
      }
    }
  }

  if (referenciaAtual) referencias.push(referenciaAtual);
  if (referencias.length) dados["referencias_profissionais"] = referencias;

  const numero = dados["n°"] || dados.numero || "";

  const agrupado = {
    nome_completo: dados.nome_completo,
    idade: dados.idade,
    pessoal: {
      fumante: dados.fumante,
      estado_civil: dados.estado_civil,
      naturalidade: dados.naturalidade,
      filhos: dados.filhos,
      saúde: dados.saúde,
      religião: dados.religião,
      cônjuge: dados.cônjuge
    },
    documentos: {
      rg: dados.rg,
      data_expedição: dados.data_expedição,
      cpf: dados.cpf,
      pis: dados.pis,
      carteira_nº: dados.carteira_nº,
      série: dados.série
    },
    endereco: {
      rua: dados.endereço,
      numero: numero,
      bairro: dados.bairro,
      cidade: dados.cidade,
      tempo_de_residência: dados.tempo_de_residência
    },
    educacao: {
      escolaridade: dados.escolaridade
    },
    profissional: {
      profissão: dados.profissão,
      cargo_desejado: dados.cargo_desejado,
      habilitação: dados.habilitação,
      categoria: dados.categoria,
      veículo_próprio: dados.veículo_próprio,
      pretensão_salarial: dados.pretensão_salarial,
      último_salário: dados.último_salário,
      disponibilidade_para_dormir: dados.disponibilidade_para_dormir,
      disponível_aos_sábados: dados.disponível_aos_sábados
    },
    obs: dados.obs,
    referencias_profissionais: dados.referencias_profissionais
  };

  const nomeFuncionario = dados.nome_completo || 'funcionario';
  const caminhoArquivoJson = path.join(__dirname, `dados-${nomeFuncionario}.json`);
  fs.writeFileSync(caminhoArquivoJson, JSON.stringify(agrupado, null, 2), 'utf8');
  console.log(`Arquivo JSON gerado: ${caminhoArquivoJson}`);
}

// Função principal para percorrer as pastas e processar arquivos
function processarArquivos() {
  const arquivos = obterArquivosNaPasta(pastaBase, 25); // Limite para 2 pastas
  arquivos.forEach(arquivo => {
    const arquivoMaisRecente = obterArquivoMaisRecente([arquivo]);
    if (arquivoMaisRecente) {
      console.log(`Processando arquivo: ${arquivoMaisRecente}`);
      processarArquivoExcel(arquivoMaisRecente);
    }
  });
}

processarArquivos();

const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const fileName = "Cadastro ( I ) Abner Carlos de Lima.xlsx";
const filePath = path.join(__dirname, fileName);

const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

let dados = {};
let referencias = [];
let referenciaAtual = null;
let modoReferencia = false;
let ultimaChave = null;

// Funções auxiliares
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

// Processa planilha
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

// === AGRUPAMENTO POR SIMILARIDADE DEPOIS DO PARSE ===
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

fs.writeFileSync("saida.json", JSON.stringify(agrupado, null, 2), "utf8");
console.log("Arquivo JSON gerado com sucesso!");

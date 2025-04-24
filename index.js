const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

// Nome do arquivo
const fileName = "Cadastro ( I ) Abner Carlos de Lima.xlsx";
const filePath = path.join(__dirname, fileName);

// Lê o arquivo Excel
const workbook = xlsx.readFile(filePath);
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

let dados = {};
let observacoesGerais = [];
let referencias = [];
let referenciaAtual = {};
let modoReferencia = false;

for (const row of rows) {
  for (const cell of row) {
    if (!cell || typeof cell !== "string") continue;

    const texto = cell.trim();
    if (!texto) continue;

    // Detecta início da seção de referências
    if (texto.toUpperCase().includes("REFERÊNCIA PROFISSIONAL")) {
      if (Object.keys(referenciaAtual).length) {
        referencias.push(referenciaAtual);
        referenciaAtual = {};
      }
      modoReferencia = true;
      continue;
    }

    if (texto.includes(":")) {
      const [rawKey, ...rawValue] = texto.split(":");
      const chave = rawKey.trim().toLowerCase().replace(/[\s\.]+/g, "_");
      const valor = rawValue.join(":").trim();

      if (modoReferencia) {
        referenciaAtual[chave] = valor;
      } else {
        dados[chave] = valor;
      }
    } else {
      // Se não tem dois pontos, trata como observação
      if (modoReferencia) {
        referenciaAtual.observacoes = (referenciaAtual.observacoes || "") + " " + texto;
      } else {
        observacoesGerais.push(texto);
      }
    }
  }
}

// Adiciona a última referência, se necessário
if (Object.keys(referenciaAtual).length) {
  referencias.push(referenciaAtual);
}

// Junta observações gerais
if (observacoesGerais.length) {
  dados["observacoes_gerais"] = observacoesGerais.join(" ");
}

// Junta referências profissionais
if (referencias.length) {
  dados["referencias_profissionais"] = referencias;
}

// Salva o resultado como JSON
fs.writeFileSync("saida.json", JSON.stringify(dados, null, 2), "utf8");
console.log("Arquivo JSON gerado com sucesso!");

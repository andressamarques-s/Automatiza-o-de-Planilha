// LISTA DE PRESENÇA E PESQUISA DE SATISFAÇÃO — BASE DE DADOS

const ABA_RESPOSTAS = "Respostas ao formulário 1";
const ABA_LIMPA     = "Base de dados (Suporte)";


// ÍNDICES DAS COLUNAS 
const IDX = {
  data_hora   : 0,   // Col A
  termo_aceite: 1,   // Col B
  nome        : 2,   // Col C
  cpf         : 3,   // Col D
  sigla_campo : 4,   // Col E
  area_servico: 33,  // Col AH
};

// Unidade: cols F a AG (índices 5 a 32) 
const IDX_UNIDADES = Array.from({ length: 28 }, (_, i) => i + 5);

// Profissional: cols AI a AL (índices 34 a 37) 
const IDX_PROFISSIONAIS = [34, 35, 36, 37];

// Campos condicionais por área de serviço
const AREAS = {
  SST: {
    servico_prestado       : 38,  // Col AM
    nivel_satisfacao_equipe: 39,  // Col AN
    nivel_satisfacao_prof  : 40,  // Col AO
    nivel_satisfacao_espera: 41,  // Col AP
    resolvido              : 42,  // Col AQ
    obs_nao_resolvido      : 43,  // Col AR
    treinamento            : 44,  // Col AS
    avaliacao_treinamento  : 45,  // Col AT
    satisfacao_treinamento : 46,  // Col AU
    avaliacao_profissional : 47,  // Col AV
  },
  AMA: {
    servico_prestado       : 48,  // Col AW
    nivel_satisfacao_equipe: 49,  // Col AX
    nivel_satisfacao_prof  : 50,  // Col AY
    nivel_satisfacao_espera: 51,  // Col AZ
    resolvido              : 52,  // Col BA
    obs_nao_resolvido      : 53,  // Col BB
    treinamento            : null,
    avaliacao_treinamento  : 54,  // Col BC
    satisfacao_treinamento : 55,  // Col BD
    avaliacao_profissional : 56,  // Col BE
  },
  GPS: {
    servico_prestado       : 57,  // Col BF
    nivel_satisfacao_equipe: 58,  // Col BG
    nivel_satisfacao_prof  : 59,  // Col BH
    nivel_satisfacao_espera: 60,  // Col BI
    resolvido              : 61,  // Col BJ
    obs_nao_resolvido      : 62,  // Col BK
    treinamento            : null,
    avaliacao_treinamento  : 63,  // Col BL
    satisfacao_treinamento : 64,  // Col BM
    avaliacao_profissional : 65,  // Col BN
  },
  SPS: {
    servico_prestado       : 66,  // Col BO
    nivel_satisfacao_equipe: 67,  // Col BP
    nivel_satisfacao_prof  : 68,  // Col BQ
    nivel_satisfacao_espera: 69,  // Col BR
    resolvido              : 70,  // Col BS
    obs_nao_resolvido      : 71,  // Col BT
    treinamento            : null,
    avaliacao_treinamento  : 72,  // Col BU
    satisfacao_treinamento : 73,  // Col BV
    avaliacao_profissional : 74,  // Col BW
  },
};

// Campos gerais 
const IDX_GERAIS = {
  nps_recomendacao: 75,  // Col BX — "Probabilidade de recomendar" (0–10) → usado para NPS real
  observacao      : 76,  // Col BY
  falar_gestor    : 77,  // Col BZ
  telefone        : 78,  // Col CA
  melhor_horario  : 79,  // Col CB
};


// CABEÇALHO DA BASE DE DADOS
const CABECALHO = [
  "data_hora",               // Data/hora formatada como texto
  "termo_aceite",
  "nome",
  "cpf",
  "sigla_campo",
  "unidade",
  "area_servico",
  "profissional",
  "servico_prestado",
  "nivel_satisfacao_equipe",
  "nivel_satisfacao_profissional",
  "nivel_satisfacao_tempo_espera",
  "resolvido",
  "obs_nao_resolvido",
  "treinamento",
  "avaliacao_treinamento",
  "satisfacao_treinamento",
  "avaliacao_profissional",
  "nps_recomendacao",        // Nota bruta (0–10) da pergunta de recomendação
  "nps_categoria",           // Promotor / Neutro / Detrator
  "observacao",
  "falar_gestor",
  "telefone",
  "melhor_horario",
];


// TRIGGER PRINCIPAL — 
function processarNovaResposta(e) {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ss.getSheetByName(ABA_RESPOSTAS);
  let   abaLimpa  = ss.getSheetByName(ABA_LIMPA);

  if (!abaLimpa) {
    abaLimpa = ss.insertSheet(ABA_LIMPA);
  }

  // Garante o cabeçalho formatado
  if (abaLimpa.getLastRow() === 0) {
    abaLimpa.appendRow(CABECALHO);
    formatarCabecalho(abaLimpa);
  }

  // Pega o índice da linha recém-enviada
  const linhaIdx = (e && e.range) ? e.range.getRow() : abaOrigem.getLastRow();
  if (!linhaIdx || linhaIdx <= 1) return;

  const linha = abaOrigem
    .getRange(linhaIdx, 1, 1, abaOrigem.getLastColumn())
    .getValues()[0];

  if (!linha || !linha[IDX.data_hora]) return;

  const linhaLimpa = montarLinhaLimpa(linha);
  if (!linhaLimpa) return;

  abaLimpa.appendRow(linhaLimpa);
  Logger.log(`✅ Linha ${linhaIdx} processada com sucesso.`);
}


// REPROCESSAR TUDO —  
function reprocessarTudo() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const abaOrigem = ss.getSheetByName(ABA_RESPOSTAS);
  let   abaLimpa  = ss.getSheetByName(ABA_LIMPA);

  if (!abaLimpa) {
    abaLimpa = ss.insertSheet(ABA_LIMPA);
  }

  const dados = abaOrigem.getDataRange().getValues();
  const linhasLimpas = [];

  for (let i = 1; i < dados.length; i++) {
    const linha = dados[i];
    if (!linha || linha.length < 10 || !linha[IDX.data_hora]) continue;
    const limpa = montarLinhaLimpa(linha);
    if (limpa) linhasLimpas.push(limpa);
  }

  // Limpa tudo e reescreve do zero
  abaLimpa.clearContents();
  abaLimpa.getRange(1, 1, 1, CABECALHO.length).setValues([CABECALHO]);

  if (linhasLimpas.length > 0) {
    abaLimpa
      .getRange(2, 1, linhasLimpas.length, CABECALHO.length)
      .setValues(linhasLimpas);
  }

  formatarCabecalho(abaLimpa);
  Logger.log(`✅ Reprocessamento concluído: ${linhasLimpas.length} linhas processadas.`);
}


// MONTAGEM DA LINHA LIMPA
function montarLinhaLimpa(linha) {
  if (!linha || !Array.isArray(linha)) return null;

  const area      = identificarArea(linha);
  const blocoArea = area ? AREAS[area] : null;

  const cpf = formatarCPF(linha[IDX.cpf]);

  // NPS real: baseado na pergunta de recomendação (col 76, índice 75)
  const notaNps    = sanitizarNumero(linha[IDX_GERAIS.nps_recomendacao]);
  const categNps   = calcularCategoriaНPS(notaNps);

  return [
    new Date(linha[IDX.data_hora]),                                              // data_hora
    linha[IDX.termo_aceite] || "",                                               // termo_aceite
    capitalizarNome(linha[IDX.nome]),                                            // nome
    cpf,                                                                         // cpf
    linha[IDX.sigla_campo] || "",                                                // sigla_campo
    getPrimeiroPreenchido(linha, IDX_UNIDADES),                                  // unidade
    normalizarArea(linha[IDX.area_servico] || area),                             // area_servico
    getPrimeiroPreenchido(linha, IDX_PROFISSIONAIS),                             // profissional
    blocoArea ? (linha[blocoArea.servico_prestado]        || "") : "",           // servico_prestado
    blocoArea ? sanitizarNumero(linha[blocoArea.nivel_satisfacao_equipe])  : "", // nivel_satisfacao_equipe
    blocoArea ? sanitizarNumero(linha[blocoArea.nivel_satisfacao_prof])    : "", // nivel_satisfacao_profissional
    blocoArea ? sanitizarNumero(linha[blocoArea.nivel_satisfacao_espera])  : "", // nivel_satisfacao_tempo_espera
    blocoArea ? (linha[blocoArea.resolvido]               || "") : "",           // resolvido
    blocoArea ? (linha[blocoArea.obs_nao_resolvido]       || "") : "",           // obs_nao_resolvido
    blocoArea && blocoArea.treinamento !== null
      ? (linha[blocoArea.treinamento] || "") : "",                               // treinamento
    blocoArea ? sanitizarNumero(linha[blocoArea.avaliacao_treinamento])    : "", // avaliacao_treinamento
    blocoArea ? sanitizarNumero(linha[blocoArea.satisfacao_treinamento])   : "", // satisfacao_treinamento
    blocoArea ? (linha[blocoArea.avaliacao_profissional]  || "") : "",           // avaliacao_profissional (texto livre)
    notaNps,                                                                     // nps_recomendacao (0–10)
    categNps,                                                                    // nps_categoria
    linha[IDX_GERAIS.observacao]       || "",                                    // observacao
    linha[IDX_GERAIS.falar_gestor]     || "",                                    // falar_gestor
    linha[IDX_GERAIS.telefone]         || "",                                    // telefone
    linha[IDX_GERAIS.melhor_horario]   || "",                                    // melhor_horario
  ];
}


// NPS — CATEGORIA POR PESSOA
// Promotor  (nota 9–10)  → "Promotor"
// Neutro    (nota 7–8)   → "Neutro"
// Detrator  (nota 0–6)   → "Detrator"

function calcularCategoriaНPS(nota) {
  if (nota === "" || nota === null || nota === undefined) return "";
  const n = Number(nota);
  if (isNaN(n)) return "";
  if (n >= 9) return "Promotor";
  if (n >= 7) return "Neutro";
  return "Detrator";
}


// IDENTIFICAR ÁREA
function identificarArea(linha) {
  if (linha[AREAS.SST.nivel_satisfacao_equipe] || linha[AREAS.SST.servico_prestado]) return "SST";
  if (linha[AREAS.AMA.nivel_satisfacao_equipe] || linha[AREAS.AMA.servico_prestado]) return "AMA";
  if (linha[AREAS.GPS.nivel_satisfacao_equipe] || linha[AREAS.GPS.servico_prestado]) return "GPS";
  if (linha[AREAS.SPS.nivel_satisfacao_equipe] || linha[AREAS.SPS.servico_prestado]) return "SPS";
  return null;
}


// UTILITÁRIOS
function getPrimeiroPreenchido(linha, indices) {
  for (const idx of indices) {
    if (linha[idx] !== undefined && linha[idx] !== null && linha[idx] !== "") {
      return linha[idx];
    }
  }
  return "";
}

function sanitizarNumero(valor) {
  if (valor === "" || valor === null || valor === undefined) return "";
  const n = Number(valor);
  return isNaN(n) ? "" : n;
}

function normalizarArea(valor) {
  if (!valor) return "";
  const v = valor.toString().trim();
  if (v.includes("AMA")) return "AMA - Assitência Médica ao Aluno";
  if (v.includes("SST")) return "SST - Saúde e Segurança do Trabalho";
  if (v.includes("GPS")) return "GPS - Gestão de Planos";
  if (v.includes("SPS")) return "SPS - Suporte de Sistemas";
  return v;
}

function formatarCPF(valor) {
  if (!valor) return "";
  const n = valor.toString().replace(/\D/g, "");
  if (n.length !== 11) return valor.toString();
  return `${n.slice(0,3)}.${n.slice(3,6)}.${n.slice(6,9)}-${n.slice(9)}`;
}

function formatarData(valor) {
  if (!valor) return "";
  const d = (valor instanceof Date) ? valor : new Date(valor);
  if (isNaN(d.getTime())) return valor.toString();
  const pad = n => String(n).padStart(2, "0");
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function capitalizarNome(nome) {
  if (!nome) return "";
  return nome.toString().toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}

function formatarCabecalho(aba) {
  const cab = aba.getRange(1, 1, 1, CABECALHO.length);
  cab.setBackground("#1a237e");
  cab.setFontColor("#ffffff");
  cab.setFontWeight("bold");
  aba.setFrozenRows(1);
  aba.autoResizeColumns(1, CABECALHO.length);


  aba.getRange("A:A").setNumberFormat("@");
}


// CONFIGURAR TRIGGER —
function configurarTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Remove triggers antigos para evitar duplicatas
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("processarNovaResposta")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
  Logger.log("✅ Trigger configurado com sucesso!");
}

/***************************************************************
 * HORALU — TIME ANDRÉ
 * Gestão SAC Magalu — Integração Genesys Cloud
 * Google Sheets + Gmail + Apps Script
 *
 * Base: script "Automação Genesys" do André (importação, base,
 * ranking, dashboard e histórico continuam iguais).
 * Novo: e-mails do HoraLu (Meta Nuvens / Meta Estrelas, ranking
 * TOP 5 / TOP 3, consolidado do time), tratamento de falta/folga,
 * modo teste, aba LOG e gatilho a cada 15 minutos.
 *
 * IMPORTANTE:
 * 1) Cole este código no Apps Script vinculado à planilha,
 *    substituindo TODO o código antigo.
 * 2) Na aba AGENTES, coluna G: confira os e-mails; coluna H: OK.
 *    Coluna C (Operação) = "Aprendiz" recebe o e-mail de aprendiz.
 * 3) Rode "Enviar teste (só para o gestor)" e confira os e-mails.
 * 4) Só depois mude ENVIAR_EMAILS_INDIVIDUAIS para true.
 ***************************************************************/

const CFG = {
  // Vazio: usa a planilha vinculada a este Apps Script.
  SPREADSHEET_ID: '',

  EMAIL_GESTOR: 'gestor@exemplo.com.br',
  EMAIL_COORDENADORA: 'coordenacao@exemplo.com.br',
  // Recebe TODOS os e-mails no modo teste.
  EMAIL_TESTE: 'teste@exemplo.com.br',

  // Nome que aparece na assinatura: "TIME ANDRÉ, o melhor time de todos!"
  NOME_TIME: 'ANDRÉ',

  // Desligado até o gestor aprovar o modo teste.
  ENVIAR_EMAILS_INDIVIDUAIS: false,
  ENVIAR_PAINEL_GESTOR: true,

  // Regra já existente na planilha: 1 atendimento a menos a cada 4,75 min ociosos.
  MINUTOS_POR_DESCONTO: 4.75,

  // HoraLu
  FATOR_ESTRELAS: 1.2,        // Meta Estrelas = meta ajustada × 1,2
  TOP_AGENTES: 5,
  TOP_APRENDIZES: 3,
  // Cancela só se faltar mais que esta fração do time nos CSVs (relatório quebrado).
  LIMITE_AUSENTES: 0.5,
  GATILHO_MINUTOS: 15,        // Valores aceitos pelo Apps Script: 1, 5, 10, 15 ou 30

  // Qualidade (relatório "Avaliações" da Genesys)
  META_CSAT: 0.92,
  META_RPC: 0.82,
  MIN_AVALIACOES: 3,          // Abaixo disso o indicador conta como na meta
  // Nota HoraLu, que ordena o ranking: pesos somam 1.
  PESOS_NOTA: { PRODUCAO: 0.7, CSAT: 0.15, RPC: 0.15 },

  // Gmail: procura anexos CSV recentes. O script só aceita CSVs com o cabeçalho da Genesys.
  // Os relatórios chegam no Gmail do André: o gatilho precisa ser instalado na conta dele.
  GMAIL_QUERY: 'newer_than:3d has:attachment filename:csv',
  TOLERANCIA_MINUTOS_ENTRE_RELATORIOS: 20,

  ABAS: {
    BASE: 'BASE_CONSOLIDADA',
    DASH: 'Dashboard',
    AGENTES: 'AGENTES',
    CONFIG: 'CONFIGURAÇÕES',
    PERFORMANCE: 'RAW_PERFORMANCE',
    OCIOSIDADE: 'RAW_OCIOSIDADE',
    RANKING: 'RANKING',
    HISTORICO: 'HISTÓRICO'
  },
  // Criada automaticamente se não existir.
  ABA_LOG: 'LOG',
  // Opcionais: se existirem, recebem o relatório de avaliações.
  ABA_AVAL_RAW: 'RAW_AVALIACOES',
  ABA_AVAL_CONSOLIDADO: 'CONSOLIDADO_AVALIACOES'
};

const STATUS_FOLGA = '⚪ Falta/Folga';

/** Menu da planilha */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('HoraLu — Time André')
    .addItem('Processar relatórios agora', 'processarRelatoriosAgora')
    .addItem('Enviar teste (só para o gestor)', 'enviarTesteParaGestor')
    .addItem('Testar leitura dos e-mails', 'testarLeituraGmail')
    .addItem('Liberar reprocessamento', 'liberarReprocessamento')
    .addSeparator()
    .addItem('Instalar execução automática (15 min)', 'instalarGatilhoHorario')
    .addItem('Remover gatilhos automáticos', 'removerGatilhos')
    .addToUi();
}

/** Função principal — pode ser executada manualmente ou pelo gatilho. */
function processarRelatoriosAgora() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Outra execução ainda está em andamento.');
  try {
  const ss = abrirPlanilha_();
  validarAbas_(ss);

  const relatorios = localizarRelatoriosGenesys_();
  if (!relatorios.performance || !relatorios.ociosidade) {
    Logger.log('Ainda não há um par de CSVs (Desempenho e Ociosidade) no Gmail. Aguardando.');
    return;
  }

  const perf = relatorios.performance;
  const oci = relatorios.ociosidade;

  const chave = [
    ss.getId(),
    perf.messageId,
    oci.messageId,
    perf.periodoInicio || '',
    perf.periodoFim || ''
  ].join('|');

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ULTIMO_PROCESSAMENTO_V2') === chave) {
    Logger.log('Este par de relatórios já foi processado. Nada a fazer.');
    return;
  }
  // Par que já falhou: não repete o erro a cada 15 minutos.
  if (props.getProperty('ULTIMO_ERRO_V2') === chave) {
    Logger.log('Este par de relatórios já falhou antes. Veja a aba LOG.');
    return;
  }

  try {
    validarCorrespondencia_(ss, perf, oci);
  } catch (e) {
    props.setProperty('ULTIMO_ERRO_V2', chave);
    registrarLog_(ss, 'ERRO', 'processamento', e.message);
    throw e;
  }

  // Atualiza abas RAW com o conteúdo integral dos CSVs.
  gravarRaw_(ss.getSheetByName(CFG.ABAS.PERFORMANCE), perf.headers, perf.rows);
  gravarRaw_(ss.getSheetByName(CFG.ABAS.OCIOSIDADE), oci.headers, oci.rows);

  // Consolida por ID de agente e atualiza a base.
  const resultado = atualizarBase_(ss, perf, oci);

  // Qualidade: CSAT e RPC do dia. Se o relatório atrasar, segue sem travar.
  const aval = relatorios.avaliacoes;
  const nomesTime = {};
  resultado.agentes.forEach(a => { nomesTime[normalizar_(a.nome)] = true; });
  const mapaAval = aval ? consolidarAvaliacoes_(aval, nomesTime, dataIso_(perf.periodoInicio)) : {};
  aplicarQualidade_(resultado.agentes, mapaAval);
  gravarQualidadeNaBase_(ss, resultado.agentes);
  gravarAvaliacoes_(ss, aval, resultado.agentes);
  if (!aval) registrarLog_(ss, 'AVISO', 'avaliações', 'Relatório de avaliações não encontrado; CSAT e RPC sem dados.');

  const ativos = resultado.agentes.filter(a => !a.ausente);

  // Atualiza ranking e painel (sem quem está de folga).
  atualizarRanking_(ss, ativos);
  atualizarDashboard_(ss, ativos, resultado.dataAnalise);

  // Registro por par de mensagens, com data/hora de atualização.
  registrarHistorico_(ss, resultado.agentes, resultado.dataAnalise);

  SpreadsheetApp.flush();

  const horaLu = montarHoraLu_(resultado.agentes);

  if (CFG.ENVIAR_EMAILS_INDIVIDUAIS) {
    enviarEmailsIndividuais_(ss, horaLu, resultado.dataAnalise, chave, false);
  }

  if (CFG.ENVIAR_PAINEL_GESTOR && props.getProperty('PAINEL_ENVIADO_V2') !== chave) {
    enviarPainelGestor_(ss, horaLu, resultado.dataAnalise, false);
    props.setProperty('PAINEL_ENVIADO_V2', chave);
  }

  props.setProperty('ULTIMO_PROCESSAMENTO_V2', chave);
  props.setProperty('ULTIMA_ATUALIZACAO', new Date().toISOString());

  registrarLog_(ss, 'OK', 'processamento',
    'Agentes: ' + resultado.agentes.length + ' | Operando: ' + ativos.length +
    ' | Período: ' + resultado.dataAnalise);
  } finally { lock.releaseLock(); }
}

/**
 * Modo teste: monta todos os e-mails com os dados atuais da planilha
 * e manda tudo só para CFG.EMAIL_TESTE. O time não recebe nada.
 */
function enviarTesteParaGestor() {
  const ss = abrirPlanilha_();
  validarAbas_(ss);
  const agentes = montarAgentesDaPlanilha_(ss);
  if (!agentes.length) {
    mostrarAlerta_('A BASE_CONSOLIDADA está vazia. Processe os relatórios antes de testar.');
    return;
  }
  const dataAnalise = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const horaLu = montarHoraLu_(agentes);
  enviarPainelGestor_(ss, horaLu, dataAnalise, true);
  enviarEmailsIndividuais_(ss, horaLu, dataAnalise, '', true);
  mostrarAlerta_('Teste enviado para ' + CFG.EMAIL_TESTE + '.\n\n' +
    'Você vai receber o consolidado e um e-mail para cada agente. A equipe não recebeu nada.');
}

function abrirPlanilha_() {
  const ss = CFG.SPREADSHEET_ID ? SpreadsheetApp.openById(CFG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Planilha não encontrada. Preencha CFG.SPREADSHEET_ID com o ID real.');
  return ss;
}

/** Localiza os CSVs no Gmail, identificando pelo cabeçalho e não pelo nome do arquivo. */
function localizarRelatoriosGenesys_() {
  const threads = GmailApp.search(CFG.GMAIL_QUERY, 0, 100);
  const performances = [];
  const ociosidades = [];
  let avaliacao = null;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const anexos = msg.getAttachments({includeInlineImages: false});
      anexos.forEach(blob => {
        const nome = (blob.getName() || '').toLowerCase();
        if (!nome.endsWith('.csv')) return;

        const texto = blob.getDataAsString('UTF-8').replace(/^﻿/, '');
        const parsed = parsearCsv_(texto);
        if (!parsed || parsed.rows.length === 0) return;

        const headersNorm = parsed.headers.map(normalizar_);
        // Relatório de avaliações: fica o mais recente que tiver linhas.
        if (headersNorm.includes('resultados com exito') && headersNorm.includes('usuarios')) {
          if (!avaliacao || msg.getDate() > avaliacao.messageDate) {
            avaliacao = { headers: parsed.headers, rows: parsed.rows,
              messageId: msg.getId(), messageDate: msg.getDate() };
          }
          return;
        }
        const temNome = headersNorm.includes('nome do agente');
        const temId = headersNorm.includes('id de agente');
        if (!temNome || !temId) return;

        const obj = {
          headers: parsed.headers,
          rows: parsed.rows,
          messageId: msg.getId(),
          messageDate: msg.getDate(),
          periodoInicio: valorPrimeiraLinha_(parsed, 'Início do intervalo'),
          periodoFim: valorPrimeiraLinha_(parsed, 'Fim do intervalo')
        };

        if (headersNorm.includes('atendidas')) {
          performances.push(obj);
        }

        if (headersNorm.includes('ocioso')) {
          ociosidades.push(obj);
        }
      });
    });
  });

  const pares = [];
  performances.forEach(p => ociosidades.forEach(o => {
    if (p.periodoInicio !== o.periodoInicio || p.periodoFim !== o.periodoFim) return;
    const delta = Math.abs(p.messageDate.getTime() - o.messageDate.getTime());
    if (delta <= CFG.TOLERANCIA_MINUTOS_ENTRE_RELATORIOS * 60000) {
      pares.push({performance:p, ociosidade:o, delta:delta,
        recente:Math.max(p.messageDate.getTime(),o.messageDate.getTime())});
    }
  }));
  pares.sort((a,b) => b.recente-a.recente || a.delta-b.delta);
  if (!pares.length) return {performance:null,ociosidade:null,avaliacoes:avaliacao};
  return Object.assign({avaliacoes: avaliacao}, pares[0]);
}

/** Lê CSV com separador vírgula, ponto e vírgula ou tabulação. */
function parsearCsv_(texto) {
  const primeiraLinha = (texto.split(/\r?\n/)[0] || '');
  const opcoes = [',', ';', '\t'];
  let separador = ',';
  let maior = -1;

  opcoes.forEach(s => {
    const qtd = primeiraLinha.split(s).length;
    if (qtd > maior) {
      maior = qtd;
      separador = s;
    }
  });

  const matriz = Utilities.parseCsv(texto, separador);
  if (!matriz || matriz.length < 2) return null;

  const headers = matriz[0].map(v => String(v || '').trim());
  const rows = matriz.slice(1)
    .filter(row => row.some(v => String(v || '').trim() !== ''))
    .map(row => {
      const ajustada = row.slice(0, headers.length);
      while (ajustada.length < headers.length) ajustada.push('');
      return ajustada.map(v => String(v == null ? '' : v).trim());
    });

  return {headers: headers, rows: rows};
}

function valorPrimeiraLinha_(parsed, titulo) {
  const idx = parsed.headers.findIndex(h => normalizar_(h) === normalizar_(titulo));
  return idx >= 0 && parsed.rows.length ? parsed.rows[0][idx] : '';
}

/** Escreve os dados brutos completos e ajusta a quantidade de colunas. */
function gravarRaw_(sheet, headers, rows) {
  sheet.clearContents();
  const dados = [headers].concat(rows);
  if (!dados.length) return;

  const totalCols = headers.length;
  const totalRows = dados.length;
  if (sheet.getMaxColumns() < totalCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), totalCols - sheet.getMaxColumns());
  }
  if (sheet.getMaxRows() < totalRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), totalRows - sheet.getMaxRows());
  }

  sheet.getRange(1, 1, totalRows, totalCols).setValues(dados);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, totalCols).setFontWeight('bold');
  sheet.autoResizeColumns(1, totalCols);
}

/** Consolida produção, TMA, ociosidade e meta ajustada. */
function atualizarBase_(ss, perf, oci) {
  const shAg = ss.getSheetByName(CFG.ABAS.AGENTES);
  const shBase = ss.getSheetByName(CFG.ABAS.BASE);
  const agValues = shAg.getDataRange().getValues();

  const perfMap = mapaPorId_(perf);
  const ociMap = mapaPorId_(oci);
  const nomesPerf = mapaPorNome_(perf);
  const nomesOci = mapaPorNome_(oci);

  const agentes = [];
  const linhasBase = [];

  // Cabeçalho da base permanece conforme o modelo atual.
  const header = [
    'Nome', 'Produção', 'Meta', '% Meta', 'TMA', 'Ociosidade',
    '% Transferência', 'Status', 'Ociosidade Min',
    'Meta Ajustada', '% Meta Ajustada', 'GAP'
  ];
  shBase.getRange(1, 1, 1, 12).setValues([header]);

  for (let r = 1; r < agValues.length; r++) {
    const id = String(agValues[r][0] || '').trim();
    const nome = String(agValues[r][1] || '').trim();
    if (!nome) continue;

    const operacao = String(agValues[r][2] || '').trim();
    const meta = numero_(agValues[r][5]);
    const email = String(agValues[r][6] || '').trim();
    const validacao = String(agValues[r][7] || '').trim();

    const pAchado = (id && perfMap[id]) || nomesPerf[normalizar_(nome)];
    const oAchado = (id && ociMap[id]) || nomesOci[normalizar_(nome)];

    // Fora dos dois relatórios: não logou hoje (falta/folga).
    if (!pAchado && !oAchado) {
      linhasBase.push([nome, '', meta, '', '', '', '', STATUS_FOLGA, '', meta, '', '']);
      agentes.push({
        id: id, nome: nome, email: email, validacaoEmail: validacao,
        operacao: operacao, ausente: true,
        producao: 0, meta: meta, metaAjustada: meta, atingimento: 0,
        tmaSeg: 0, ociosidadeSeg: 0, ociosidadeMin: 0,
        transferidas: 0, percTransferencia: 0, gap: 0, status: STATUS_FOLGA
      });
      continue;
    }

    const p = pAchado || {};
    const o = oAchado || {};

    const producao = numero_(p.atendidas);
    const tmaSeg = duracaoSegundos_(p.tratamentoMedio);
    const ociSeg = duracaoSegundos_(o.ocioso);
    const ociMin = ociSeg / 60;
    const transferidas = numero_(p.transferidas);
    // Proxy operacional: transferidas / atendidas. Valide com a definição de taxa Genesys.
    const percTransferencia = producao > 0 ? transferidas / producao : 0;

    // Regra da planilha: 1 atendimento descontado a cada 4,75 min ociosos.
    const desconto = Math.floor(ociMin / CFG.MINUTOS_POR_DESCONTO);
    const metaAjustada = Math.max(0, meta - desconto);
    const atingimento = metaAjustada > 0 ? producao / metaAjustada : (producao > 0 ? 1 : 0);
    const atingimentoBase = meta > 0 ? producao / meta : 0;
    const gap = producao - metaAjustada;

    let status = '🔴 Crítico';
    if (atingimento >= 1) status = '🟢 Meta Batida';
    else if (atingimento >= 0.8) status = '🟡 Atenção';

    const tmaDias = tmaSeg / 86400;
    const ociDias = ociSeg / 86400;

    linhasBase.push([
      nome, producao, meta, atingimentoBase, tmaDias, ociDias,
      percTransferencia, status, ociMin, metaAjustada, atingimento, gap
    ]);

    agentes.push({
      id: id,
      nome: nome,
      email: email,
      validacaoEmail: validacao,
      operacao: operacao,
      ausente: false,
      producao: producao,
      meta: meta,
      metaAjustada: metaAjustada,
      atingimento: atingimento,
      tmaSeg: tmaSeg,
      ociosidadeSeg: ociSeg,
      ociosidadeMin: ociMin,
      transferidas: transferidas,
      percTransferencia: percTransferencia,
      gap: gap,
      status: status
    });
  }

  // Limpa valores antigos sem remover formatação.
  const maxLinhas = Math.max(shBase.getMaxRows() - 1, linhasBase.length);
  if (maxLinhas > 0) {
    shBase.getRange(2, 1, maxLinhas, 12).clearContent();
  }
  if (linhasBase.length) {
    shBase.getRange(2, 1, linhasBase.length, 12).setValues(linhasBase);
    shBase.getRange(2, 4, linhasBase.length, 1).setNumberFormat('0.00%');
    shBase.getRange(2, 5, linhasBase.length, 2).setNumberFormat('[h]:mm:ss');
    shBase.getRange(2, 7, linhasBase.length, 1).setNumberFormat('0.00%');
    shBase.getRange(2, 9, linhasBase.length, 1).setNumberFormat('0.00');
    shBase.getRange(2, 11, linhasBase.length, 1).setNumberFormat('0.00%');
  }

  const dataAnalise = perf.periodoInicio || Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'
  );

  return {agentes: agentes, dataAnalise: dataAnalise};
}

/**
 * Reconstrói a lista de agentes a partir da BASE_CONSOLIDADA e da aba AGENTES.
 * Usado pelo modo teste, sem precisar de um relatório novo no Gmail.
 */
function montarAgentesDaPlanilha_(ss) {
  const cadastro = {};
  const agValues = ss.getSheetByName(CFG.ABAS.AGENTES).getDataRange().getValues();
  for (let r = 1; r < agValues.length; r++) {
    const nome = String(agValues[r][1] || '').trim();
    if (!nome) continue;
    cadastro[normalizar_(nome)] = {
      id: String(agValues[r][0] || '').trim(),
      operacao: String(agValues[r][2] || '').trim(),
      email: String(agValues[r][6] || '').trim(),
      validacao: String(agValues[r][7] || '').trim()
    };
  }

  const shBase = ss.getSheetByName(CFG.ABAS.BASE);
  const ultima = shBase.getLastRow();
  if (ultima < 2) return [];
  const dados = shBase.getRange(2, 1, ultima - 1, 19).getValues();

  const agentes = [];
  dados.forEach(l => {
    const nome = String(l[0] || '').trim();
    if (!nome) return;
    const c = cadastro[normalizar_(nome)] || {};
    const ausente = l[1] === '' || l[1] === null;
    const meta = numero_(l[2]);
    const metaAjustada = ausente ? meta : numero_(l[9]);
    const producao = ausente ? 0 : numero_(l[1]);
    agentes.push({
      id: c.id || '',
      nome: nome,
      email: c.email || '',
      validacaoEmail: c.validacao || '',
      operacao: c.operacao || '',
      ausente: ausente,
      producao: producao,
      meta: meta,
      metaAjustada: metaAjustada,
      atingimento: ausente ? 0 : numero_(l[10]),
      gap: ausente ? 0 : numero_(l[11]),
      status: ausente ? STATUS_FOLGA : String(l[7] || ''),
      csat: l[12] === '' ? null : numero_(l[12]),
      csatResp: numero_(l[13]),
      csatPromotor: Math.round(numero_(l[12]) * numero_(l[13])),
      rpc: l[14] === '' ? null : numero_(l[14]),
      rpcResp: numero_(l[15]),
      rpcSim: Math.round(numero_(l[14]) * numero_(l[15])),
      rpcNao: numero_(l[15]) - Math.round(numero_(l[14]) * numero_(l[15])),
      csatDetrator: numero_(l[17]),
      csatNeutro: numero_(l[18])
    });
  });
  agentes.forEach(a => calcularNota_(a, CFG.PESOS_NOTA));
  return agentes;
}

/** Índices do relatório de desempenho por ID de agente. */
function mapaPorId_(relatorio) {
  const h = relatorio.headers.map(normalizar_);
  const idx = {
    id: h.indexOf('id de agente'),
    nome: h.indexOf('nome do agente'),
    atendidas: h.indexOf('atendidas'),
    tratamentoMedio: h.indexOf('tratamento medio'),
    transferidas: h.indexOf('transferidas'),
    ocioso: h.indexOf('ocioso')
  };
  const mapa = {};
  relatorio.rows.forEach(row => {
    const id = idx.id >= 0 ? String(row[idx.id] || '').trim() : '';
    if (!id) return;
    mapa[id] = {
      nome: idx.nome >= 0 ? row[idx.nome] : '',
      atendidas: idx.atendidas >= 0 ? row[idx.atendidas] : 0,
      tratamentoMedio: idx.tratamentoMedio >= 0 ? row[idx.tratamentoMedio] : '',
      transferidas: idx.transferidas >= 0 ? row[idx.transferidas] : 0,
      ocioso: idx.ocioso >= 0 ? row[idx.ocioso] : ''
    };
  });
  return mapa;
}

/** Mapa por nome para contingência quando o ID não estiver cadastrado. */
function mapaPorNome_(relatorio) {
  const h = relatorio.headers.map(normalizar_);
  const idxNome = h.indexOf('nome do agente');
  const idxId = h.indexOf('id de agente');
  const idxAtendidas = h.indexOf('atendidas');
  const idxTratMedio = h.indexOf('tratamento medio');
  const idxTransferidas = h.indexOf('transferidas');
  const idxOcioso = h.indexOf('ocioso');

  const mapa = {};
  relatorio.rows.forEach(row => {
    const nome = idxNome >= 0 ? normalizar_(row[idxNome]) : '';
    if (!nome) return;
    mapa[nome] = {
      id: idxId >= 0 ? row[idxId] : '',
      atendidas: idxAtendidas >= 0 ? row[idxAtendidas] : 0,
      tratamentoMedio: idxTratMedio >= 0 ? row[idxTratMedio] : '',
      transferidas: idxTransferidas >= 0 ? row[idxTransferidas] : 0,
      ocioso: idxOcioso >= 0 ? row[idxOcioso] : ''
    };
  });
  return mapa;
}

function atualizarRanking_(ss, agentes) {
  const sh = ss.getSheetByName(CFG.ABAS.RANKING);
  sh.getRange(1, 1, 1, 5).setValues([['Posição', 'Agente', '% Meta Ajustada', 'GAP', 'Nota HoraLu']]);

  const ordenados = ordenarRanking_(agentes);

  const dados = ordenados.map((a, i) => [
    i + 1, a.nome, a.atingimento, a.gap, a.nota
  ]);

  const limpar = Math.max(sh.getMaxRows() - 1, dados.length);
  if (limpar > 0) sh.getRange(2, 1, limpar, 5).clearContent();
  if (dados.length) {
    sh.getRange(2, 1, dados.length, 5).setValues(dados);
    sh.getRange(2, 3, dados.length, 1).setNumberFormat('0.00%');
    sh.getRange(2, 5, dados.length, 1).setNumberFormat('0.0%');
  }
}

function atualizarDashboard_(ss, agentes, dataAnalise) {
  const sh = ss.getSheetByName(CFG.ABAS.DASH);
  const total = agentes.length;
  const producao = agentes.reduce((s, a) => s + a.producao, 0);
  const metaAj = agentes.reduce((s, a) => s + a.metaAjustada, 0);
  const bateram = agentes.filter(a => a.atingimento >= 1).length;
  const atencao = agentes.filter(a => a.atingimento >= 0.8 && a.atingimento < 1).length;
  const criticos = agentes.filter(a => a.atingimento < 0.8).length;
  const pctGeral = metaAj > 0 ? producao / metaAj : 0;
  const gapMedio = agentes.length
    ? agentes.reduce((s, a) => s + a.gap, 0) / agentes.length : 0;

  sh.getRange('B1').setValue('Período: ' + dataAnalise + ' | Atualizado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM HH:mm'));
  sh.getRange('B2:B6').setValues([
    [total], [bateram], [atencao], [criticos], [0]
  ]); // B6: categoria intermediária descontinuada
  sh.getRange('E2:E5').setValues([
    [total ? producao / total : 0],
    [total ? metaAj / total : 0],
    [pctGeral],
    [gapMedio]
  ]);
  sh.getRange('E4').setNumberFormat('0.00%');

  const top = ordenarRanking_(agentes).slice(0, 5);
  const crit = agentes.filter(a => a.atingimento < 0.8)
    .sort((a, b) => a.atingimento - b.atingimento).slice(0, 5);

  sh.getRange('G2:H6').clearContent();
  sh.getRange('J2:K6').clearContent();
  if (top.length) {
    sh.getRange(2, 7, top.length, 2).setValues(
      top.map(a => [a.nome, a.atingimento])
    );
    sh.getRange(2, 8, top.length, 1).setNumberFormat('0.00%');
  }
  if (crit.length) {
    sh.getRange(2, 10, crit.length, 2).setValues(
      crit.map(a => [a.nome, a.atingimento])
    );
    sh.getRange(2, 11, crit.length, 1).setNumberFormat('0.00%');
  }
}

function registrarHistorico_(ss, agentes, dataAnalise) {
  const sh = ss.getSheetByName(CFG.ABAS.HISTORICO);
  const linhas = agentes.map(a => [
    new Date(), a.nome, a.ausente ? '' : a.producao, a.meta, a.metaAjustada,
    a.ausente ? '' : a.atingimento, a.ausente ? '' : a.gap,
    a.ausente ? '' : a.tmaSeg / 86400,
    a.ausente ? '' : a.ociosidadeSeg / 86400, a.status
  ]);
  if (!linhas.length) return;

  const inicio = sh.getLastRow() + 1;
  sh.getRange(inicio, 1, linhas.length, 10).setValues(linhas);
  sh.getRange(inicio, 1, linhas.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange(inicio, 6, linhas.length, 1).setNumberFormat('0.00%');
  sh.getRange(inicio, 8, linhas.length, 2).setNumberFormat('[h]:mm:ss');
}

/* ==============================================================
 * HORALU — qualidade (CSAT e RPC)
 * O relatório "Avaliações" da Genesys traz uma linha por ligação. O
 * resultado da pesquisa vem como etiquetas na coluna "Resultados com
 * êxito": "[CSAT SAC] - RPC SIM/NÃO" e "[CSAT SAC] - PROMOTOR/NEUTRO/DETRATOR".
 * ============================================================== */

/** "29/10/26 12:07" ou "29/10/2026" -> "2026-10-29" (vazio se não reconhecer). */
function dataIso_(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const m = String(valor || '').match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return '';
  const ano = m[3].length === 2 ? '20' + m[3] : m[3];
  return ano + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

/**
 * Conta as avaliações por agente (chave = nome normalizado).
 * Só considera ligações do dia analisado e de agentes da aba AGENTES.
 * Em ligação transferida, a avaliação vai para o último agente da lista.
 */
function consolidarAvaliacoes_(aval, nomesTime, diaIso) {
  const h = aval.headers.map(normalizar_);
  const idxConversa = h.indexOf('id de conversa');
  const idxUsuarios = h.indexOf('usuarios');
  const idxData = h.indexOf('data');
  const idxResultado = h.indexOf('resultados com exito');
  const mapa = {};
  if (idxUsuarios < 0 || idxResultado < 0) return mapa;

  const vistas = {};
  aval.rows.forEach(row => {
    if (diaIso && idxData >= 0 && dataIso_(row[idxData]) !== diaIso) return;
    const conversa = idxConversa >= 0 ? row[idxConversa] : '';
    if (conversa && vistas[conversa]) return;
    if (conversa) vistas[conversa] = true;

    const etiquetas = String(row[idxResultado] || '').split(';').map(normalizar_)
      .filter(t => t.indexOf('csat') >= 0);
    if (!etiquetas.length) return;

    const usuarios = String(row[idxUsuarios] || '').split(';').map(normalizar_).filter(Boolean);
    const agente = usuarios[usuarios.length - 1];
    if (!agente || !nomesTime[agente]) return;

    const c = mapa[agente] || (mapa[agente] = { promotor: 0, neutro: 0, detrator: 0, rpcSim: 0, rpcNao: 0 });
    etiquetas.forEach(t => {
      if (/rpc sim/.test(t)) c.rpcSim++;
      else if (/rpc nao/.test(t)) c.rpcNao++;
      else if (/promotor/.test(t)) c.promotor++;
      else if (/neutro/.test(t)) c.neutro++;
      else if (/detrator/.test(t)) c.detrator++;
    });
  });
  return mapa;
}

/** Acrescenta CSAT, RPC e a nota HoraLu em cada agente. */
function aplicarQualidade_(agentes, mapa) {
  const pesos = CFG.PESOS_NOTA;
  agentes.forEach(a => {
    const c = mapa[normalizar_(a.nome)] || { promotor: 0, neutro: 0, detrator: 0, rpcSim: 0, rpcNao: 0 };
    a.csatResp = c.promotor + c.neutro + c.detrator;
    a.csatPromotor = c.promotor;
    a.csatNeutro = c.neutro;
    a.csatDetrator = c.detrator;
    a.csat = a.csatResp ? c.promotor / a.csatResp : null;
    a.rpcResp = c.rpcSim + c.rpcNao;
    a.rpcSim = c.rpcSim;
    a.rpcNao = c.rpcNao;
    a.rpc = a.rpcResp ? c.rpcSim / a.rpcResp : null;
    calcularNota_(a, pesos);
  });
}

/**
 * Nota HoraLu (0 a 1,2): produção pesa mais e qualidade completa.
 * Com menos de CFG.MIN_AVALIACOES respostas, o indicador conta como na meta.
 */
function calcularNota_(a, pesos) {
  const fator = (valor, resp, meta) =>
    resp >= CFG.MIN_AVALIACOES && valor !== null ? Math.min(CFG.FATOR_ESTRELAS, valor / meta) : 1;
  a.fatorCsat = fator(a.csat, a.csatResp || 0, CFG.META_CSAT);
  a.fatorRpc = fator(a.rpc, a.rpcResp || 0, CFG.META_RPC);
  a.notaQualidade = pesos.CSAT * a.fatorCsat + pesos.RPC * a.fatorRpc;
  a.nota = a.ausente ? 0 : pesos.PRODUCAO * Math.min(CFG.FATOR_ESTRELAS, a.atingimento) + a.notaQualidade;
}

/** Grava o CSV cru e a tabela por agente nas abas de avaliações (se existirem). */
function gravarAvaliacoes_(ss, aval, agentes) {
  const shRaw = ss.getSheetByName(CFG.ABA_AVAL_RAW);
  if (shRaw && aval) gravarRaw_(shRaw, aval.headers, aval.rows);

  const sh = ss.getSheetByName(CFG.ABA_AVAL_CONSOLIDADO);
  if (!sh) return;
  const linhas = [['Agente', 'Respostas CSAT', 'Promotores', 'Neutros', 'Detratores', 'CSAT',
    'Respostas RPC', 'RPC Sim', 'RPC Não', 'RPC', 'Nota HoraLu']]
    .concat(agentes.map(a => [
      a.nome, a.csatResp || 0, a.csatPromotor || 0, a.csatNeutro || 0, a.csatDetrator || 0,
      a.csat === null ? '' : a.csat,
      a.rpcResp || 0, a.rpcSim || 0, a.rpcNao || 0, a.rpc === null ? '' : a.rpc,
      a.ausente ? '' : a.nota
    ]));
  sh.clearContents();
  sh.getRange(1, 1, linhas.length, 11).setValues(linhas);
  sh.getRange(1, 1, 1, 11).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (linhas.length > 1) {
    sh.getRange(2, 6, linhas.length - 1, 1).setNumberFormat('0.0%');
    sh.getRange(2, 10, linhas.length - 1, 2).setNumberFormat('0.0%');
  }
}

/** Colunas extras da BASE_CONSOLIDADA (M a S), usadas também pelo modo teste. */
function gravarQualidadeNaBase_(ss, agentes) {
  const sh = ss.getSheetByName(CFG.ABAS.BASE);
  sh.getRange(1, 13, 1, 7).setValues([['CSAT', 'Respostas CSAT', 'RPC', 'Respostas RPC', 'Nota HoraLu', 'Detratores', 'Neutros']]);
  const dados = agentes.map(a => [
    a.csat === null ? '' : a.csat, a.csatResp || 0,
    a.rpc === null ? '' : a.rpc, a.rpcResp || 0,
    a.ausente ? '' : a.nota, a.csatDetrator || 0, a.csatNeutro || 0
  ]);
  if (!dados.length) return;
  sh.getRange(2, 13, dados.length, 7).setValues(dados);
  sh.getRange(2, 13, dados.length, 1).setNumberFormat('0.0%');
  sh.getRange(2, 15, dados.length, 1).setNumberFormat('0.0%');
  sh.getRange(2, 17, dados.length, 1).setNumberFormat('0.0%');
}

/** "12 promotores · 1 neutro · 2 detratores" (o neutro só aparece se houver). */
function detalheCsat_(x) {
  const plural = (n, um, varios) => n + ' ' + (n === 1 ? um : varios);
  const partes = ["<span style='color:" + COR.verde + ";font-weight:bold;'>" + plural(x.csatPromotor || 0, 'promotor', 'promotores') + "</span>"];
  if (x.csatNeutro) partes.push(plural(x.csatNeutro, 'neutro', 'neutros'));
  partes.push("<span style='color:" + COR.vermelho + ";font-weight:bold;'>" + plural(x.csatDetrator || 0, 'detrator', 'detratores') + "</span>");
  return partes.join(' · ');
}

/** "10 sim · 2 não" */
function detalheRpc_(x) {
  return "<span style='color:" + COR.verde + ";font-weight:bold;'>" + (x.rpcSim || 0) + " sim</span> · " +
    "<span style='color:" + COR.vermelho + ";font-weight:bold;'>" + (x.rpcNao || 0) + " não</span>";
}

/** Visual de um indicador de qualidade: cor, texto do valor e detalhe. */
function visualQualidade_(valor, resp, meta) {
  if (!resp) return { cor: COR.cinza, valor: '–', detalhe: 'sem avaliações hoje', fundo: '#F1F3F5' };
  const n = resp + (resp === 1 ? ' avaliação' : ' avaliações');
  if (resp < CFG.MIN_AVALIACOES) {
    return { cor: COR.cinza, valor: pct_(valor * 100, 0), detalhe: 'poucas avaliações (' + n + ')', fundo: '#F1F3F5' };
  }
  const cor = valor >= meta ? COR.verde : valor >= meta - 0.1 ? COR.laranja : COR.vermelho;
  const fundo = valor >= meta ? '#E3F7EC' : valor >= meta - 0.1 ? '#FFF3E0' : '#FFE8EF';
  return { cor: cor, valor: pct_(valor * 100, 0), detalhe: n + ' · meta ' + pct_(meta * 100, 0), fundo: fundo };
}

/* ==============================================================
 * HORALU — cálculo de metas Nuvens/Estrelas e ranking
 * ============================================================== */

/** Ordena pela nota HoraLu (produção + CSAT + RPC); empate desempata pela produção. */
function ordenarRanking_(lista) {
  const nota = a => a.nota === undefined ? a.atingimento : a.nota;
  return lista.slice().sort((a, b) =>
    nota(b) - nota(a) || b.producao - a.producao
  );
}

/**
 * Atendimentos que faltam para a nota do agente passar a do alvo, mantendo
 * a qualidade atual. null quando só produção não basta (precisa de CSAT/RPC).
 */
function faltaParaPassar_(agente, alvo) {
  if (agente.nota === undefined) {
    if (agente.metaAjustada > 0) {
      return Math.max(1, Math.floor(alvo.atingimento * agente.metaAjustada) + 1 - agente.producao);
    }
    return Math.max(1, alvo.producao - agente.producao + 1);
  }
  const atingNecessario = (alvo.nota - agente.notaQualidade) / CFG.PESOS_NOTA.PRODUCAO;
  if (atingNecessario >= CFG.FATOR_ESTRELAS) return null;
  if (agente.metaAjustada <= 0) return Math.max(1, alvo.producao - agente.producao + 1);
  return Math.max(1, Math.floor(atingNecessario * agente.metaAjustada) + 1 - agente.producao);
}

/**
 * Monta tudo o que os e-mails do HoraLu precisam: status Nuvens/Estrelas,
 * rankings separados de agentes e aprendizes e o consolidado do time.
 */
function montarHoraLu_(agentes) {
  const pessoas = agentes.map(a => {
    const ehAprendiz = normalizar_(a.operacao) === 'aprendiz';
    const p = Object.assign({}, a, { ehAprendiz: ehAprendiz });
    if (a.ausente) {
      p.statusTabela = 'Falta/Folga';
      p.corTabela = '#e8e8e8';
      p.faixa = 'folga';
      return p;
    }

    p.metaNuvens = a.metaAjustada;
    p.metaEstrelas = Math.round(a.metaAjustada * CFG.FATOR_ESTRELAS);
    const restamNuvens = p.metaNuvens - a.producao;
    const restamEstrelas = p.metaEstrelas - a.producao;
    p.restamNuvens = Math.max(0, restamNuvens);
    p.restamEstrelas = Math.max(0, restamEstrelas);
    p.entregou = restamNuvens <= 0;
    p.faixa = restamEstrelas <= 0 ? 'dupla' : restamNuvens <= 0 ? 'nuvens'
      : a.atingimento >= 0.8 ? 'atencao' : 'critico';

    const nuvem = ehAprendiz ? '' : ' ☁️';
    const estrela = ehAprendiz ? '' : ' ⭐';
    if (restamNuvens > 0) {
      p.fraseStatus = 'para a Meta Nuvens' + nuvem + ' (100%) faltam ' + restamNuvens +
        ' atendimentos e para a Meta Estrelas' + estrela + ' (120%) faltam ' + restamEstrelas + ' atendimentos' +
        (ehAprendiz ? '. Acreditamos muito no seu potencial!' : '.');
      p.statusTabela = ehAprendiz ? 'Faltam ' + restamNuvens + ' Nuvens' : 'Faltam ' + restamNuvens + ' ☁️';
      p.corTabela = '#ffebee';
    } else if (restamEstrelas > 0) {
      p.fraseStatus = 'parabéns! Você BATEU a Meta Nuvens' + nuvem + '! Agora o seu desafio é a Meta Estrelas' +
        estrela + ' (120%), faltam apenas ' + restamEstrelas + ' atendimentos' +
        (ehAprendiz ? '. Continue com esse foco maravilhoso!' : '.');
      p.statusTabela = ehAprendiz ? 'Nuvens OK! Faltam ' + restamEstrelas + ' Estrelas' : 'Nuvens OK! Faltam ' + restamEstrelas + ' ⭐';
      p.corTabela = '#e8f5e9';
    } else {
      p.fraseStatus = ehAprendiz
        ? 'sensacional! Você superou todas as expectativas e BATEU as duas metas do dia: Nuvens e Estrelas! Você é um verdadeiro talento da nossa equipe!'
        : 'sensacional! Você superou todas as expectativas e BATEU as duas metas do dia: Nuvens ☁️ e Estrelas ⭐! Você é gigante!';
      p.statusTabela = ehAprendiz ? 'Dupla Meta Batida!' : 'Dupla Meta! ☁️⭐';
      p.corTabela = '#e1f5fe';
    }
    return p;
  });

  const ativos = pessoas.filter(p => !p.ausente);
  const agentesAtivos = ordenarRanking_(ativos.filter(p => !p.ehAprendiz));
  const aprendizesAtivos = ordenarRanking_(ativos.filter(p => p.ehAprendiz));

  calcularRankingTexto_(agentesAtivos, CFG.TOP_AGENTES, false);
  calcularRankingTexto_(aprendizesAtivos, CFG.TOP_APRENDIZES, true);

  const prod = ativos.reduce((s, p) => s + p.producao, 0);
  const meta = ativos.reduce((s, p) => s + p.metaNuvens, 0);
  const sim = ativos.filter(p => p.entregou).length;
  const somar = campo => ativos.reduce((s, p) => s + (p[campo] || 0), 0);
  const csatResp = somar('csatResp'), rpcResp = somar('rpcResp');

  return {
    agentesAtivos: agentesAtivos,
    aprendizesAtivos: aprendizesAtivos,
    agentesAusentes: pessoas.filter(p => p.ausente && !p.ehAprendiz),
    aprendizesAusentes: pessoas.filter(p => p.ausente && p.ehAprendiz),
    time: {
      producao: prod,
      meta: meta,
      pct: meta > 0 ? prod / meta * 100 : 0,
      sim: sim,
      nao: ativos.length - sim,
      pctEntrega: ativos.length ? sim / ativos.length * 100 : 0,
      csatResp: csatResp,
      csat: csatResp ? somar('csatPromotor') / csatResp : null,
      csatPromotor: somar('csatPromotor'),
      csatNeutro: somar('csatNeutro'),
      csatDetrator: somar('csatDetrator'),
      rpcSim: somar('rpcSim'),
      rpcNao: somar('rpcNao'),
      rpcResp: rpcResp,
      rpc: rpcResp ? somar('rpcSim') / rpcResp : null
    }
  };
}

/** Preenche posição e as frases do card de ranking de cada pessoa (lista já ordenada). */
function calcularRankingTexto_(lista, top, ehAprendiz) {
  lista.forEach((p, i) => {
    p.posicao = i + 1;
    p.noTop = p.posicao <= top;
    if (p.posicao === 1) {
      p.rankingTitulo = ehAprendiz ? 'Você é o destaque dos aprendizes!' : 'Você lidera o ranking do time!';
      p.rankingTexto = ehAprendiz
        ? 'Que orgulho! Continue dando o seu melhor.'
        : 'Desempenho fantástico. Agora é segurar a liderança!';
    } else if (p.noTop) {
      const acima = lista[i - 1];
      p.faltaSubir = faltaParaPassar_(p, acima);
      p.rankingTitulo = 'Você está no TOP ' + top + (ehAprendiz ? ' dos aprendizes!' : ' do time!');
      p.rankingTexto = p.faltaSubir === null
        ? 'Para subir para o <b>' + (p.posicao - 1) + 'º lugar</b>, capriche na qualidade: CSAT e RPC também contam na nota.'
        : 'Faltam <b>' + p.faltaSubir + '</b> atendimentos para você subir para o <b>' + (p.posicao - 1) + 'º lugar</b>.';
    } else {
      const corte = lista[top - 1];
      p.faltaSubir = faltaParaPassar_(p, corte);
      p.rankingTitulo = 'Rumo ao TOP ' + top + '!';
      p.rankingTexto = (p.faltaSubir === null
        ? 'Para entrar no <b>TOP ' + top + '</b>, junte produção e qualidade: CSAT e RPC também contam na nota. '
        : 'Faltam <b>' + p.faltaSubir + '</b> atendimentos para você entrar no <b>TOP ' + top + '</b>. ') +
        (ehAprendiz ? 'Acelera que dá!' : 'Você consegue, acelera!');
    }
  });
}

function primeiroNome_(nome) {
  const p = String(nome || '').trim().split(/\s+/)[0] || '';
  return p.charAt(0).toUpperCase() + p.slice(1).toLowerCase();
}

function pct_(valor, casas) {
  return valor.toFixed(casas).replace('.', ',') + '%';
}

/* ==============================================================
 * HORALU — e-mails (visual Magalu)
 * HTML de e-mail: só tabelas e estilos inline, que é o que o Gmail
 * Emojis fora do básico (fogo, medalhas) vão como entidade HTML (&#x...;):
 * o GmailApp troca esses caracteres por "?" no corpo do e-mail.
 * e o Outlook respeitam. Sem JavaScript (nenhum leitor de e-mail roda).
 * ============================================================== */

const COR = {
  azul: '#0086FF', azulEscuro: '#0050C8', marinho: '#0A2A66',
  rosa: '#E5007E', amarelo: '#FFC300', laranja: '#FF8A00',
  verde: '#00A651', vermelho: '#E5004F', cinza: '#9AA5B1',
  fundo: '#EEF3FA', texto: '#1F2A44', suave: '#5B6B82', linha: '#E3E9F2'
};

/** Visual de cada faixa de desempenho. */
const FAIXAS = {
  dupla:   { rotulo: 'Dupla Meta',  cor: COR.azul,     fundo: '#E3F1FF', icone: '☁️⭐' },
  nuvens:  { rotulo: 'Nuvens OK',   cor: COR.verde,    fundo: '#E3F7EC', icone: '☁️' },
  atencao: { rotulo: 'Atenção',     cor: COR.laranja,  fundo: '#FFF3E0', icone: '⚡' },
  critico: { rotulo: 'Crítico',     cor: COR.vermelho, fundo: '#FFE8EF', icone: '&#x1F525;' },
  folga:   { rotulo: 'Falta/Folga', cor: COR.cinza,    fundo: '#F1F3F5', icone: '' }
};

/**
 * E-mail individual. Só envia se a coluna H da aba AGENTES estiver como OK.
 * No modo teste, manda tudo para CFG.EMAIL_TESTE e ignora o controle de reenvio.
 */
function enviarEmailsIndividuais_(ss, horaLu, dataAnalise, chave, modoTeste) {
  const props = PropertiesService.getScriptProperties();
  const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
  const pessoas = horaLu.agentesAtivos.concat(horaLu.aprendizesAtivos);

  pessoas.forEach(p => {
    if (!modoTeste && !emailAprovado_(p.email, p.validacaoEmail)) return;
    const controle = 'AGENTE_V2_' + Utilities.base64EncodeWebSafe(p.id || p.nome).slice(0,60);
    if (!modoTeste && props.getProperty(controle) === chave) return;

    const grupo = p.ehAprendiz ? horaLu.aprendizesAtivos : horaLu.agentesAtivos;
    const html = htmlIndividual_(p, grupo, horaLu.time, hora);
    const destino = modoTeste ? CFG.EMAIL_TESTE : p.email;
    const assunto = modoTeste
      ? '[TESTE VISÃO ' + (p.ehAprendiz ? 'APRENDIZ' : 'AGENTE') + '] Como chegará para: ' + p.nome
      : assuntoIndividual_(p, hora);

    try {
      GmailApp.sendEmail(destino, assunto, 'Seu acompanhamento de desempenho está disponível em formato HTML.', {
        htmlBody: html,
        name: 'HoraLu — Time ' + CFG.NOME_TIME
      });
      if (!modoTeste) props.setProperty(controle, chave);
    } catch (e) {
      registrarLog_(ss, 'ERRO', destino, 'Falha no envio para ' + p.nome + ': ' + e.message);
    }
  });
}

/** Assunto que já conta a novidade na caixa de entrada. */
function assuntoIndividual_(p, hora) {
  const nome = primeiroNome_(p.nome);
  const pos = p.posicao + 'º';
  if (p.faixa === 'dupla') return 'Dupla Meta, ' + nome + '! Você está em ' + pos + ' | HoraLu ' + hora;
  if (p.faixa === 'nuvens') return nome + ', Nuvens batida! Faltam ' + p.restamEstrelas + ' para as Estrelas | HoraLu ' + hora;
  return nome + ', faltam ' + p.restamNuvens + ' para a Meta Nuvens. Você está em ' + pos + ' | HoraLu ' + hora;
}

/* ---------- blocos visuais ---------- */

/** Faixa colorida da marca Magalu. */
function faixaMagalu_() {
  const cores = [COR.amarelo, COR.laranja, COR.vermelho, COR.rosa, '#8E44EF', COR.azul, COR.verde];
  return "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;'><tr>" +
    cores.map(c => "<td height='6' style='height:6px;line-height:6px;font-size:0;background:" + c + ";'>&nbsp;</td>").join('') +
    "</tr></table>";
}

/** Barra de progresso de 0 a 120%, com as marcas de Nuvens (100%) e Estrelas (120%). */
function barraMetas_(atingimento, cor, comMarcas) {
  const preenchido = Math.max(2, Math.min(100, Math.round(atingimento / CFG.FATOR_ESTRELAS * 100)));
  const marcaNuvens = Math.round(100 / CFG.FATOR_ESTRELAS);
  let h = "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:separate;background:#DCE5F0;border-radius:10px;'>" +
    "<tr><td width='" + preenchido + "%' height='16' style='height:16px;font-size:0;line-height:16px;background:" + cor + ";border-radius:10px;'>&nbsp;</td>" +
    (preenchido < 100 ? "<td style='font-size:0;line-height:16px;'>&nbsp;</td>" : '') + "</tr></table>";
  if (comMarcas) {
    h += "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;margin-top:4px;'><tr>" +
      "<td width='" + marcaNuvens + "%' align='right' style='font-size:11px;color:" + COR.suave + ";border-right:2px solid " + COR.azul + ";padding-right:4px;'>☁️ 100%</td>" +
      "<td align='right' style='font-size:11px;color:" + COR.suave + ";'>⭐ " + Math.round(CFG.FATOR_ESTRELAS * 100) + "%</td>" +
      "</tr></table>";
  }
  return h;
}

/** Barrinha compacta para tabelas. */
function miniBarra_(atingimento, cor) {
  const largura = Math.max(3, Math.min(100, Math.round(atingimento / CFG.FATOR_ESTRELAS * 100)));
  return "<table role='presentation' width='90' cellpadding='0' cellspacing='0' style='border-collapse:separate;background:#DCE5F0;border-radius:6px;'>" +
    "<tr><td width='" + largura + "%' height='8' style='height:8px;font-size:0;line-height:8px;background:" + cor + ";border-radius:6px;'>&nbsp;</td>" +
    (largura < 100 ? "<td style='font-size:0;line-height:8px;'>&nbsp;</td>" : '') + "</tr></table>";
}

function pill_(texto, cor, fundo) {
  return "<span style='display:inline-block;padding:3px 10px;border-radius:12px;font-size:11px;font-weight:bold;color:" + cor +
    ";background:" + fundo + ";white-space:nowrap;'>" + texto + "</span>";
}

/** Cartão de número grande (KPI). */
function tile_(rotulo, valor, cor, detalhe) {
  return "<td valign='top' style='padding:6px;'>" +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:#FFFFFF;border:1px solid " + COR.linha + ";border-top:4px solid " + cor + ";border-radius:10px;'>" +
    "<tr><td style='padding:12px 12px 2px 12px;font-size:11px;color:" + COR.suave + ";text-transform:uppercase;letter-spacing:0.5px;'>" + rotulo + "</td></tr>" +
    "<tr><td style='padding:0 12px;font-size:24px;font-weight:bold;color:" + cor + ";'>" + valor + "</td></tr>" +
    "<tr><td style='padding:2px 12px 12px 12px;font-size:11px;color:" + COR.suave + ";'>" + (detalhe || '&nbsp;') + "</td></tr>" +
    "</table></td>";
}

function cabecalho_(titulo, subtitulo, corDe, corAte) {
  return faixaMagalu_() +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:" + corDe + ";background-image:linear-gradient(135deg," + corDe + " 0%," + corAte + " 100%);'>" +
    "<tr><td style='padding:22px 26px;'>" +
    "<div style='font-size:12px;color:#FFFFFF;opacity:0.85;letter-spacing:1px;text-transform:uppercase;font-weight:bold;'>HoraLu · Time " + escaparHtml_(CFG.NOME_TIME) + "</div>" +
    "<div style='font-size:24px;color:#FFFFFF;font-weight:bold;margin-top:4px;'>" + titulo + "</div>" +
    "<div style='font-size:13px;color:#FFFFFF;opacity:0.9;margin-top:4px;'>" + subtitulo + "</div>" +
    "</td></tr></table>";
}

function moldura_(conteudo, largura) {
  return "<html><head><meta charset='UTF-8'><meta name='viewport' content='width=device-width,initial-scale=1'></head>" +
    "<body style='margin:0;padding:0;background:" + COR.fundo + ";font-family:Arial,Helvetica,sans-serif;color:" + COR.texto + ";'>" +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:" + COR.fundo + ";'><tr><td align='center' style='padding:20px 10px;'>" +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='max-width:" + largura + "px;background:#FFFFFF;border-radius:14px;overflow:hidden;border:1px solid " + COR.linha + ";'>" +
    "<tr><td>" + conteudo + "</td></tr></table>" +
    "</td></tr></table></body></html>";
}

function assinatura_() {
  return "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:#F6F9FD;border-top:1px solid " + COR.linha + ";'>" +
    "<tr><td align='center' style='padding:16px;font-size:14px;font-weight:bold;color:" + COR.azulEscuro + ";'>" +
    "<span style='color:" + COR.rosa + ";'>❤</span> TIME " + escaparHtml_(CFG.NOME_TIME) + ", o melhor time de todos! <span style='color:" + COR.rosa + ";'>❤</span></td></tr>" +
    "<tr><td align='center' style='padding:0 16px 14px 16px;font-size:11px;color:" + COR.cinza + ";'>SAC Magazine Luiza · HoraLu, atualizado automaticamente com os dados da Genesys Cloud</td></tr>" +
    "</table>" + faixaMagalu_();
}

/** Nome + sobrenome, para caber nas tabelas e diferenciar xarás. */
function nomeCurto_(nome) {
  const partes = String(nome || '').trim().split(/\s+/);
  const cap = s => s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
  return partes.length > 1 ? cap(partes[0]) + ' ' + cap(partes[partes.length - 1]) : cap(partes[0] || '');
}

/* ---------- e-mail individual (agente e aprendiz) ---------- */

function htmlIndividual_(p, grupo, time, hora) {
  const aprendiz = p.ehAprendiz;
  const f = FAIXAS[p.faixa];
  const destaque = aprendiz ? COR.rosa : COR.azul;
  const nuvem = aprendiz ? '' : '☁️ ';
  const estrela = aprendiz ? '' : '⭐ ';

  let manchete, apoio;
  if (p.faixa === 'dupla') {
    manchete = 'Dupla Meta batida! ' + (aprendiz ? '' : '&#x1F3C6;');
    apoio = aprendiz ? 'Você superou as duas metas do dia. Você é um verdadeiro talento da nossa equipe!'
      : 'Você superou as duas metas do dia. Você é gigante!';
  } else if (p.faixa === 'nuvens') {
    manchete = nuvem + 'Meta Nuvens batida!';
    apoio = 'Agora o desafio é a Meta Estrelas: faltam só <b>' + p.restamEstrelas + '</b> atendimentos.';
  } else {
    manchete = 'Faltam ' + p.restamNuvens + ' para a ' + nuvem + 'Meta Nuvens';
    apoio = (aprendiz ? 'Acreditamos muito no seu potencial! ' : 'Ajuste o ritmo e vamos juntos! ') +
      'Para a Meta Estrelas faltam <b>' + p.restamEstrelas + '</b>.';
  }

  const titulo = aprendiz ? 'E aí, ' + escaparHtml_(primeiroNome_(p.nome)) + '! Hora do show!'
    : 'Olá, ' + escaparHtml_(primeiroNome_(p.nome)) + '! Foco nas métricas!';

  // Card principal: % + barra até as Estrelas
  const progresso =
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:" + f.fundo + ";border-radius:12px;'>" +
    "<tr><td style='padding:18px 20px 6px 20px;'>" +
      "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr>" +
      "<td valign='middle' style='font-size:44px;font-weight:bold;color:" + f.cor + ";line-height:1;'>" + pct_(p.atingimento * 100, 0) + "</td>" +
      "<td valign='middle' align='right'>" + pill_((aprendiz ? '' : f.icone + ' ') + f.rotulo, '#FFFFFF', f.cor) + "</td>" +
      "</tr></table>" +
      "<div style='font-size:12px;color:" + COR.suave + ";margin-top:2px;'>da sua meta ajustada de hoje</div>" +
    "</td></tr>" +
    "<tr><td style='padding:10px 20px 4px 20px;'>" + barraMetas_(p.atingimento, f.cor, !aprendiz) + "</td></tr>" +
    "<tr><td style='padding:12px 20px 18px 20px;'>" +
      "<div style='font-size:18px;font-weight:bold;color:" + COR.texto + ";'>" + manchete + "</div>" +
      "<div style='font-size:14px;color:" + COR.suave + ";margin-top:4px;line-height:1.5;'>" + apoio + "</div>" +
    "</td></tr></table>";

  // Três números
  const numeros =
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='margin-top:14px;'><tr>" +
    tile_('Produção', p.producao, COR.marinho, 'atendimentos até agora') +
    tile_(nuvem + 'Nuvens', p.metaNuvens, COR.azul, p.restamNuvens ? 'faltam ' + p.restamNuvens : 'batida ✔') +
    tile_(estrela + 'Estrelas', p.metaEstrelas, aprendiz ? COR.rosa : '#C98A00', p.restamEstrelas ? 'faltam ' + p.restamEstrelas : 'batida ✔') +
    "</tr></table>";

  // Qualidade: CSAT e RPC pessoais
  const vCsat = visualQualidade_(p.csat, p.csatResp || 0, CFG.META_CSAT);
  const vRpc = visualQualidade_(p.rpc, p.rpcResp || 0, CFG.META_RPC);
  const qualidade =
    "<div style='font-size:11px;color:" + COR.suave + ";text-transform:uppercase;letter-spacing:0.5px;margin:18px 6px 0 6px;'>Sua qualidade hoje</div>" +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='margin-top:4px;'><tr>" +
    tile_('CSAT', vCsat.valor, vCsat.cor, vCsat.detalhe + (p.csatResp ? '<br>' + detalheCsat_(p) : '')) +
    tile_('RPC', vRpc.valor, vRpc.cor, vRpc.detalhe + (p.rpcResp ? '<br>' + detalheRpc_(p) : '')) +
    "</tr></table>";

  // Ranking: posição + mini placar
  const top = aprendiz ? CFG.TOP_APRENDIZES : CFG.TOP_AGENTES;
  const medalha = ['&#x1F947;', '&#x1F948;', '&#x1F949;'];
  let placar = '';
  grupo.slice(0, top).forEach((g, i) => {
    const eu = g === p;
    placar += linhaPlacar_(i < 3 && !aprendiz ? medalha[i] : (i + 1) + 'º', g, eu, destaque);
  });
  if (!p.noTop) {
    placar += "<tr><td colspan='3' align='center' style='padding:2px;font-size:12px;color:" + COR.cinza + ";'>· · ·</td></tr>" +
      linhaPlacar_(p.posicao + 'º', p, true, destaque);
  }
  const ranking =
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='margin-top:16px;border:1px solid " + COR.linha + ";border-radius:12px;'>" +
    "<tr><td style='padding:16px 18px 6px 18px;'>" +
      "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr>" +
      "<td valign='middle' width='72' align='center' style='background:" + destaque + ";border-radius:12px;color:#FFFFFF;font-size:28px;font-weight:bold;padding:10px 0;'>" + p.posicao + "º</td>" +
      "<td valign='middle' style='padding-left:14px;'>" +
        "<div style='font-size:16px;font-weight:bold;color:" + COR.texto + ";'>" + p.rankingTitulo + "</div>" +
        "<div style='font-size:13px;color:" + COR.suave + ";margin-top:3px;line-height:1.4;'>" + p.rankingTexto + "</div>" +
        "<div style='font-size:12px;color:" + COR.suave + ";margin-top:6px;'>Sua nota HoraLu: <b style='color:" + destaque + ";'>" + Math.round(p.nota * 100) + " pts</b> · " +
        Math.round(CFG.PESOS_NOTA.PRODUCAO * 100) + "% produção + " + Math.round(CFG.PESOS_NOTA.CSAT * 100) + "% CSAT + " + Math.round(CFG.PESOS_NOTA.RPC * 100) + "% RPC</div>" +
      "</td></tr></table>" +
    "</td></tr>" +
    "<tr><td style='padding:8px 12px 14px 12px;'>" +
      "<div style='font-size:11px;color:" + COR.suave + ";text-transform:uppercase;letter-spacing:0.5px;padding:0 6px 6px 6px;'>TOP " + top + (aprendiz ? ' aprendizes' : ' do time') + " agora · dados dos colegas são privados</div>" +
      "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'>" + placar + "</table>" +
    "</td></tr></table>";

  // Time
  const corTime = time.pct >= 100 ? COR.verde : COR.laranja;
  const equipe =
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='margin-top:16px;background:#F6F9FD;border-radius:12px;'>" +
    "<tr><td style='padding:16px 18px 6px 18px;'>" +
      "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr>" +
      "<td style='font-size:14px;font-weight:bold;color:" + COR.texto + ";'>Resultado do time</td>" +
      "<td align='right' style='font-size:20px;font-weight:bold;color:" + corTime + ";'>" + pct_(time.pct, 1) + "</td>" +
      "</tr></table></td></tr>" +
    "<tr><td style='padding:4px 18px;'>" + barraMetas_(time.pct / 100, corTime, false) + "</td></tr>" +
    "<tr><td style='padding:8px 18px 4px 18px;font-size:13px;color:" + COR.suave + ";'>" +
      "<b style='color:" + COR.verde + ";'>" + time.sim + "</b> pessoas já bateram a Meta Nuvens e <b style='color:" + COR.vermelho + ";'>" + time.nao + "</b> ainda estão na busca. Bora fechar forte!" +
    "</td></tr>" +
    "<tr><td style='padding:4px 18px 16px 18px;font-size:13px;color:" + COR.suave + ";'>" +
      "CSAT do time: <b style='color:" + visualQualidade_(time.csat, time.csatResp, CFG.META_CSAT).cor + ";'>" + visualQualidade_(time.csat, time.csatResp, CFG.META_CSAT).valor + "</b>" +
      " &nbsp;·&nbsp; RPC do time: <b style='color:" + visualQualidade_(time.rpc, time.rpcResp, CFG.META_RPC).cor + ";'>" + visualQualidade_(time.rpc, time.rpcResp, CFG.META_RPC).valor + "</b>" +
    "</td></tr></table>";

  const corpo = cabecalho_(titulo, 'Seus números atualizados às ' + hora,
      aprendiz ? COR.azul : COR.azulEscuro, aprendiz ? COR.rosa : COR.azul) +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr><td style='padding:22px 22px 26px 22px;'>" +
    progresso + numeros + qualidade + ranking + equipe +
    "</td></tr></table>" + assinatura_();

  return moldura_(corpo, 600);
}

/**
 * Linha do placar no e-mail individual. Por regra interna, o agente só vê os
 * próprios dados: dos colegas aparece apenas a posição, com nome e % mascarados.
 * O mascaramento é feito aqui, sem colocar o dado real no HTML (um "borrão" por
 * CSS deixaria o dado no código-fonte do e-mail).
 */
function linhaPlacar_(posicao, g, eu, destaque) {
  const fundo = eu ? '#FFF7D6' : '#FFFFFF';
  const bloco = (largura, altura) => "<span style='display:inline-block;width:" + largura + "px;height:" + altura +
    "px;border-radius:" + altura + "px;background:#DCE3EC;vertical-align:middle;'></span>";
  const larguras = [118, 96, 132, 104, 88];
  const nome = eu ? 'Você' : bloco(larguras[(g.posicao || 0) % larguras.length], 12);
  const valor = eu ? Math.round((g.nota === undefined ? g.atingimento : g.nota) * 100) + ' pts' : bloco(36, 12);
  const corValor = eu ? FAIXAS[g.faixa].cor : COR.cinza;
  return "<tr style='background:" + fundo + ";'>" +
    "<td width='44' align='center' style='padding:7px 4px;font-size:14px;font-weight:bold;color:" + COR.texto + ";border-bottom:1px solid " + COR.linha + ";'>" + posicao + "</td>" +
    "<td style='padding:7px 4px;font-size:13px;color:" + COR.texto + ";border-bottom:1px solid " + COR.linha + ";" + (eu ? "font-weight:bold;color:" + destaque + ";" : '') + "'>" + nome + "</td>" +
    "<td width='70' align='right' style='padding:7px 8px;font-size:13px;font-weight:bold;color:" + corValor + ";border-bottom:1px solid " + COR.linha + ";'>" + valor + "</td>" +
    "</tr>";
}

/* ---------- consolidado da gestão ---------- */

/** Consolidado para o gestor (coordenação em cópia). No teste, só para CFG.EMAIL_TESTE. */
function enviarPainelGestor_(ss, horaLu, dataAnalise, modoTeste) {
  if (!CFG.EMAIL_GESTOR || CFG.EMAIL_GESTOR.indexOf('@') < 0) {
    throw new Error('Preencha CFG.EMAIL_GESTOR com o e-mail corporativo do gestor.');
  }
  const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm');
  const html = htmlGestor_(horaLu, dataAnalise, hora, ss.getUrl());
  const t = horaLu.time;
  const assunto = (modoTeste ? '[TESTE GESTOR] ' : '') + 'HoraLu Gestão ' + hora + ' | Time em ' + pct_(t.pct, 1) +
    ' | ' + t.sim + ' de ' + (t.sim + t.nao) + ' na Meta Nuvens';
  const destino = modoTeste ? CFG.EMAIL_TESTE : CFG.EMAIL_GESTOR;
  const opcoes = { htmlBody: html, name: 'HoraLu — Gestão' };
  if (!modoTeste && CFG.EMAIL_COORDENADORA) opcoes.cc = CFG.EMAIL_COORDENADORA;

  try {
    GmailApp.sendEmail(destino, assunto, 'Resumo de desempenho do time.', opcoes);
  } catch (e) {
    registrarLog_(ss, 'ERRO', destino, 'Falha no envio do consolidado: ' + e.message);
  }
}

function htmlGestor_(horaLu, dataAnalise, hora, urlPlanilha) {
  const t = horaLu.time;
  const ativos = horaLu.agentesAtivos.concat(horaLu.aprendizesAtivos);
  const ausentes = horaLu.agentesAusentes.concat(horaLu.aprendizesAusentes);
  const conta = faixa => ativos.filter(a => a.faixa === faixa).length;
  const qtd = { dupla: conta('dupla'), nuvens: conta('nuvens'), atencao: conta('atencao'), critico: conta('critico'), folga: ausentes.length };
  const corTime = t.pct >= 100 ? COR.verde : t.pct >= 80 ? COR.laranja : COR.vermelho;
  const faltamTime = Math.max(0, t.meta - t.producao);

  // KPIs
  const vCsatTime = visualQualidade_(t.csat, t.csatResp || 0, CFG.META_CSAT);
  const vRpcTime = visualQualidade_(t.rpc, t.rpcResp || 0, CFG.META_RPC);
  const kpis =
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr>" +
    tile_('Atingimento do time', pct_(t.pct, 1), corTime, t.producao + ' de ' + t.meta + ' atendimentos') +
    tile_('Bateram Nuvens', t.sim + '/' + (t.sim + t.nao), COR.verde, pct_(t.pctEntrega, 0) + ' do time operando') +
    "</tr><tr>" +
    tile_('Faltam para a meta', faltamTime, COR.azul, 'atendimentos no consolidado') +
    tile_('Em risco', qtd.critico, COR.vermelho, 'abaixo de 80% · ' + qtd.folga + ' falta/folga') +
    "</tr><tr>" +
    tile_('CSAT do time', vCsatTime.valor, vCsatTime.cor, (t.csatResp || 0) + ' avaliações · meta ' + pct_(CFG.META_CSAT * 100, 0) + '<br>' + detalheCsat_(t)) +
    tile_('RPC do time', vRpcTime.valor, vRpcTime.cor, (t.rpcResp || 0) + ' avaliações · meta ' + pct_(CFG.META_RPC * 100, 0) + '<br>' + detalheRpc_(t)) +
    "</tr></table>";

  // Distribuição empilhada
  const totalDist = ativos.length + ausentes.length || 1;
  const ordemFaixas = ['dupla', 'nuvens', 'atencao', 'critico', 'folga'];
  const segmentos = ordemFaixas.filter(k => qtd[k] > 0).map(k =>
    "<td width='" + Math.max(4, Math.round(qtd[k] / totalDist * 100)) + "%' height='22' align='center' style='background:" + FAIXAS[k].cor +
    ";color:#FFFFFF;font-size:12px;font-weight:bold;'>" + qtd[k] + "</td>").join('');
  const legenda = ordemFaixas.map(k =>
    "<td style='padding:6px 8px 0 0;font-size:11px;color:" + COR.suave + ";white-space:nowrap;'>" +
    "<span style='display:inline-block;width:9px;height:9px;border-radius:5px;background:" + FAIXAS[k].cor + ";'></span> " +
    FAIXAS[k].rotulo + "</td>").join('');
  const distribuicao =
    secao_('Como o time está distribuído') +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:separate;border-radius:8px;overflow:hidden;'><tr>" + segmentos + "</tr></table>" +
    "<table role='presentation' cellpadding='0' cellspacing='0'><tr>" + legenda + "</tr></table>";

  // Pódio
  const podio = horaLu.agentesAtivos.slice(0, 3);
  const ordemPodio = [1, 0, 2].filter(i => podio[i]);
  const alturas = {0: 86, 1: 64, 2: 50};
  const coresPodio = {0: COR.amarelo, 1: '#B8C4D6', 2: '#E0A36B'};
  const podioHtml = podio.length ? secao_('Pódio da hora') +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr>" +
    ordemPodio.map(i =>
      "<td valign='bottom' align='center' width='33%' style='padding:0 6px;'>" +
      "<div style='font-size:13px;font-weight:bold;color:" + COR.texto + ";'>" + escaparHtml_(nomeCurto_(podio[i].nome)) + "</div>" +
      "<div style='font-size:12px;color:" + COR.suave + ";margin:2px 0 6px 0;'>" + Math.round(podio[i].nota * 100) + " pts · " + pct_(podio[i].atingimento * 100, 0) + " da meta</div>" +
      "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr><td height='" + alturas[i] + "' align='center' valign='top' style='height:" + alturas[i] + "px;background:" + coresPodio[i] +
      ";border-radius:10px 10px 0 0;font-size:22px;font-weight:bold;color:#FFFFFF;padding-top:8px;'>" + (i + 1) + "º</td></tr></table>" +
      "</td>").join('') +
    "</tr></table>" : '';

  // Ranking completo
  const tabela = (titulo, ativosG, ausentesG) => {
    if (!ativosG.length && !ausentesG.length) return '';
    let h = secao_(titulo) +
      "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;font-size:12px;'>" +
      "<tr style='background:" + COR.marinho + ";color:#FFFFFF;'>" +
      ["#", "Agente", "Prod.", "Faltam", "% meta aj.", "CSAT", "RPC", "Nota", "Status"].map((c, i) =>
        "<th align='" + (i === 1 || i === 8 ? 'left' : 'center') + "' style='padding:8px 5px;font-size:11px;font-weight:bold;'>" + c + "</th>").join('') + "</tr>";
    const celulaQualidade = (valor, resp, meta, bom, ruim, letraBom, letraRuim) => {
      const v = visualQualidade_(valor, resp || 0, meta);
      return "<td align='center' style='padding:7px 5px;border-bottom:1px solid " + COR.linha + ";white-space:nowrap;'>" +
        "<div style='font-weight:bold;color:" + v.cor + ";'>" + v.valor + "</div>" +
        "<div style='font-size:10px;color:" + COR.cinza + ";'>" + (resp || 0) + " aval.</div>" +
        (resp ? "<div style='font-size:10px;'><span style='color:" + COR.verde + ";font-weight:bold;'>" + (bom || 0) + letraBom + "</span> · " +
          "<span style='color:" + COR.vermelho + ";font-weight:bold;'>" + (ruim || 0) + letraRuim + "</span></div>" : '') +
        "</td>";
    };
    ativosG.forEach((a, i) => {
      const f = FAIXAS[a.faixa];
      const zebra = i % 2 ? '#F8FAFD' : '#FFFFFF';
      h += "<tr style='background:" + zebra + ";'>" +
        "<td align='center' style='padding:7px 6px;font-weight:bold;border-bottom:1px solid " + COR.linha + ";'>" + a.posicao + "º</td>" +
        "<td style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";border-left:4px solid " + f.cor + ";'>" + escaparHtml_(nomeCurto_(a.nome)) + "</td>" +
        "<td align='center' style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";'>" + a.producao + "</td>" +
        "<td align='center' style='padding:7px 6px;font-weight:bold;color:" + (a.restamNuvens ? COR.vermelho : COR.verde) + ";border-bottom:1px solid " + COR.linha + ";'>" + (a.restamNuvens || '✔') + "</td>" +
        "<td align='center' style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";'>" +
          "<div style='font-weight:bold;color:" + f.cor + ";margin-bottom:3px;'>" + pct_(a.atingimento * 100, 1) + "</div>" + miniBarra_(a.atingimento, f.cor) + "</td>" +
        celulaQualidade(a.csat, a.csatResp, CFG.META_CSAT, a.csatPromotor, a.csatDetrator, 'P', 'D') +
        celulaQualidade(a.rpc, a.rpcResp, CFG.META_RPC, a.rpcSim, a.rpcNao, 'S', 'N') +
        "<td align='center' style='padding:7px 5px;font-weight:bold;font-size:13px;color:" + COR.marinho + ";border-bottom:1px solid " + COR.linha + ";'>" + Math.round(a.nota * 100) + "</td>" +
        "<td style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";'>" + pill_(f.rotulo, f.cor, f.fundo) + "</td></tr>";
    });
    ausentesG.forEach(a => {
      h += "<tr style='background:#F6F7F9;color:" + COR.cinza + ";'>" +
        "<td align='center' style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";'>–</td>" +
        "<td style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";border-left:4px solid " + COR.cinza + ";'>" + escaparHtml_(nomeCurto_(a.nome)) + "</td>" +
        "<td colspan='6' align='center' style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";'>sem registro na Genesys</td>" +
        "<td style='padding:7px 6px;border-bottom:1px solid " + COR.linha + ";'>" + pill_('Falta/Folga', COR.cinza, FAIXAS.folga.fundo) + "</td></tr>";
    });
    return h + "</table>";
  };

  // Onde agir: quem está abaixo de 80%, do mais distante para o mais perto
  const risco = ativos.filter(a => a.faixa === 'critico').sort((a, b) => b.restamNuvens - a.restamNuvens);
  const agir = risco.length ? secao_('Onde agir agora') +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:#FFF5F8;border:1px solid #FFD0DE;border-radius:10px;'>" +
    risco.map(a =>
      "<tr><td style='padding:8px 14px;font-size:13px;color:" + COR.texto + ";border-bottom:1px solid #FFE1EA;'>&#x1F525; <b>" + escaparHtml_(nomeCurto_(a.nome)) + "</b></td>" +
      "<td align='right' style='padding:8px 14px;font-size:12px;color:" + COR.suave + ";border-bottom:1px solid #FFE1EA;'>" +
      pct_(a.atingimento * 100, 0) + " · faltam <b style='color:" + COR.vermelho + ";'>" + a.restamNuvens + "</b> para Nuvens</td></tr>").join('') +
    "</table>" : '';

  // Qualidade abaixo da meta (só com avaliações suficientes)
  const abaixoQualidade = ativos.filter(a =>
    ((a.csatResp || 0) >= CFG.MIN_AVALIACOES && a.csat < CFG.META_CSAT) ||
    ((a.rpcResp || 0) >= CFG.MIN_AVALIACOES && a.rpc < CFG.META_RPC));
  const agirQualidade = abaixoQualidade.length ? secao_('Qualidade abaixo da meta') +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' style='background:#FFF8EC;border:1px solid #FFE0B2;border-radius:10px;'>" +
    abaixoQualidade.map(a => {
      const partes = [];
      if ((a.csatResp || 0) >= CFG.MIN_AVALIACOES && a.csat < CFG.META_CSAT) partes.push('CSAT <b style="color:' + COR.vermelho + ';">' + pct_(a.csat * 100, 0) + '</b> (' + a.csatPromotor + ' prom. · ' + a.csatDetrator + ' detr.)');
      if ((a.rpcResp || 0) >= CFG.MIN_AVALIACOES && a.rpc < CFG.META_RPC) partes.push('RPC <b style="color:' + COR.vermelho + ';">' + pct_(a.rpc * 100, 0) + '</b> (' + a.rpcSim + ' sim · ' + a.rpcNao + ' não)');
      return "<tr><td style='padding:8px 14px;font-size:13px;color:" + COR.texto + ";border-bottom:1px solid #FFEBC8;'><b>" + escaparHtml_(nomeCurto_(a.nome)) + "</b></td>" +
        "<td align='right' style='padding:8px 14px;font-size:12px;color:" + COR.suave + ";border-bottom:1px solid #FFEBC8;'>" + partes.join(' · ') + "</td></tr>";
    }).join('') +
    "</table>" : '';

  const botao = urlPlanilha ?
    "<table role='presentation' cellpadding='0' cellspacing='0' align='center' style='margin:24px auto 4px auto;'><tr>" +
    "<td align='center' style='background:" + COR.azul + ";border-radius:24px;'>" +
    "<a href='" + urlPlanilha + "' style='display:inline-block;padding:12px 28px;font-size:14px;font-weight:bold;color:#FFFFFF;text-decoration:none;'>Abrir painel completo na planilha →</a>" +
    "</td></tr></table>" : '';

  const corpo = cabecalho_('Painel da Gestão', 'Atualização das ' + hora + ' · Período ' + escaparHtml_(dataAnalise), COR.marinho, COR.azul) +
    "<table role='presentation' width='100%' cellpadding='0' cellspacing='0'><tr><td style='padding:18px 18px 26px 18px;'>" +
    kpis + distribuicao + podioHtml + agir + agirQualidade +
    tabela('Ranking completo · agentes', horaLu.agentesAtivos, horaLu.agentesAusentes) +
    tabela('Ranking · aprendizes', horaLu.aprendizesAtivos, horaLu.aprendizesAusentes) +
    "<p style='font-size:11px;color:" + COR.cinza + ";margin:12px 0 0 0;'>Ranking pela nota HoraLu = " +
    Math.round(CFG.PESOS_NOTA.PRODUCAO * 100) + "% produção (% da meta ajustada, até " + Math.round(CFG.FATOR_ESTRELAS * 100) + "%) + " +
    Math.round(CFG.PESOS_NOTA.CSAT * 100) + "% CSAT + " + Math.round(CFG.PESOS_NOTA.RPC * 100) + "% RPC (cada um em relação à meta; com menos de " +
    CFG.MIN_AVALIACOES + " avaliações conta como na meta). CSAT = promotores ÷ respostas (P = promotores, D = detratores) · RPC = sim ÷ respostas (S = sim, N = não). Meta ajustada = meta − 1 atendimento a cada " +
    String(CFG.MINUTOS_POR_DESCONTO).replace('.', ',') + " min ociosos). Nuvens = 100% da meta ajustada · Estrelas = " +
    Math.round(CFG.FATOR_ESTRELAS * 100) + "% · Atenção = 80% a 99% · Crítico &lt; 80%.</p>" +
    botao +
    "</td></tr></table>" + assinatura_();

  return moldura_(corpo, 720);
}

function secao_(titulo) {
  return "<div style='font-size:13px;font-weight:bold;color:" + COR.marinho + ";text-transform:uppercase;letter-spacing:0.6px;margin:24px 0 10px 0;'>" +
    "<span style='display:inline-block;width:4px;height:14px;background:" + COR.azul + ";border-radius:2px;vertical-align:middle;margin-right:6px;'></span>" + titulo + "</div>";
}

/* ==============================================================
 * Utilitários
 * ============================================================== */

function testarLeituraGmail() {
  const rel = localizarRelatoriosGenesys_();
  Logger.log('Desempenho encontrado: ' + Boolean(rel.performance));
  Logger.log('Ociosidade encontrada: ' + Boolean(rel.ociosidade));
  Logger.log('Hora do teste: ' + new Date());
  if (rel.performance) {
    Logger.log('CSV desempenho: ' + rel.performance.headers.join(' | '));
    Logger.log('Linhas desempenho: ' + rel.performance.rows.length);
    Logger.log('Início desempenho: ' + rel.performance.periodoInicio);
    Logger.log('Fim desempenho: ' + rel.performance.periodoFim);
    Logger.log('Recebido desempenho: ' + rel.performance.messageDate);
    Logger.log('ID e-mail desempenho: ' + rel.performance.messageId);
  }
  if (rel.ociosidade) {
    Logger.log('CSV ociosidade: ' + rel.ociosidade.headers.join(' | '));
    Logger.log('Linhas ociosidade: ' + rel.ociosidade.rows.length);
    Logger.log('Início ociosidade: ' + rel.ociosidade.periodoInicio);
    Logger.log('Fim ociosidade: ' + rel.ociosidade.periodoFim);
    Logger.log('Recebido ociosidade: ' + rel.ociosidade.messageDate);
    Logger.log('ID e-mail ociosidade: ' + rel.ociosidade.messageId);
  }
}

function instalarGatilhoHorario() {
  removerGatilhos();
  ScriptApp.newTrigger('processarRelatoriosAgora')
    .timeBased()
    .everyMinutes(CFG.GATILHO_MINUTOS)
    .create();
  mostrarAlerta_(
    'Gatilho instalado. O HoraLu verificará o Gmail a cada ' + CFG.GATILHO_MINUTOS +
    ' minutos e envia os e-mails assim que um relatório novo da Genesys chegar.'
  );
}

function removerGatilhos() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'processarRelatoriosAgora') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function validarAbas_(ss) {
  Object.keys(CFG.ABAS).forEach(k => {
    const nome = CFG.ABAS[k];
    if (!ss.getSheetByName(nome)) {
      throw new Error('A aba obrigatória não foi encontrada: ' + nome);
    }
  });
}

function emailAprovado_(email, validacao) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email || '') &&
    normalizar_(validacao) === 'ok';
}

/** Registra uma linha na aba LOG (cria a aba se não existir). */
function registrarLog_(ss, tipo, destino, mensagem) {
  try {
    let sh = ss.getSheetByName(CFG.ABA_LOG);
    if (!sh) {
      sh = ss.insertSheet(CFG.ABA_LOG);
      sh.appendRow(['Data/hora', 'Tipo', 'Destino', 'Mensagem']);
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), tipo, destino, mensagem]);
  } catch (e) {
    Logger.log('Falha ao registrar log: ' + e.message);
  }
  Logger.log(tipo + ' | ' + destino + ' | ' + mensagem);
}

/** Alerta na tela quando há alguém com a planilha aberta; no gatilho, vai para o Logger. */
function mostrarAlerta_(mensagem) {
  try { SpreadsheetApp.getUi().alert(mensagem); }
  catch (e) { Logger.log(mensagem.replace(/\n/g, ' ')); }
}

function duracaoSegundos_(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;
  if (typeof valor === 'number') return valor * 86400;

  const s = String(valor).trim();
  const m = s.match(/^(-)?(\d+):([0-5]?\d):([0-5]?\d)(?:[.,](\d+))?$/);
  if (m) {
    const sinal = m[1] ? -1 : 1;
    const h = Number(m[2]), min = Number(m[3]), sec = Number(m[4]);
    const frac = m[5] ? Number('0.' + m[5]) : 0;
    return sinal * (h * 3600 + min * 60 + sec + frac);
  }
  const n = Number(s.replace(',', '.'));
  return isFinite(n) ? n * 86400 : 0;
}

function numero_(valor) {
  if (typeof valor === 'number') return valor;
  const s = String(valor == null ? '' : valor).trim().replace(',', '.');
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function normalizar_(valor) {
  return String(valor == null ? '' : valor)
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ');
}

function escaparHtml_(valor) {
  return String(valor == null ? '' : valor)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function liberarReprocessamento() {
  const propriedades = PropertiesService.getScriptProperties();

  propriedades.deleteProperty('ULTIMO_PROCESSAMENTO');
  propriedades.deleteProperty('ULTIMO_PROCESSAMENTO_V2');
  propriedades.deleteProperty('ULTIMO_ERRO_V2');
  // O consolidado pode ser reenviado; os controles dos e-mails individuais
  // permanecem para impedir reenvio acidental ao time.
  propriedades.deleteProperty('PAINEL_ENVIADO_V2');

  Logger.log('Controle de processamento reiniciado com sucesso!');
}

/**
 * Confere se os agentes da aba AGENTES estão nos dois relatórios.
 * Quem não aparece vira Falta/Folga. Só cancela se faltar mais que
 * CFG.LIMITE_AUSENTES do time (sinal de relatório quebrado ou de outro time).
 */
function validarCorrespondencia_(ss, perf, oci) {
  const dados = ss.getSheetByName(CFG.ABAS.AGENTES).getDataRange().getValues();
  const pId = mapaPorId_(perf), oId = mapaPorId_(oci);
  const pNome = mapaPorNome_(perf), oNome = mapaPorNome_(oci);
  const faltantes = [];
  let total = 0;
  for (let i=1;i<dados.length;i++) {
    const id = String(dados[i][0] || '').trim();
    const nome = String(dados[i][1] || '').trim();
    if (!nome) continue;
    total++;
    const chaveNome = normalizar_(nome);
    if (!((id && pId[id]) || pNome[chaveNome]) &&
        !((id && oId[id]) || oNome[chaveNome])) {
      faltantes.push(nome);
    }
  }
  if (total && faltantes.length > total * CFG.LIMITE_AUSENTES) throw new Error(
    'Importação cancelada: ' + faltantes.length + ' de ' + total +
    ' agentes não estão nos CSVs (relatório incompleto ou de outro time?): ' +
    faltantes.join('; ')
  );
  return faltantes;
}

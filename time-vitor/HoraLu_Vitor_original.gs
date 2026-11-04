/***************************************************************
 * HORALU — VERSÃO ORIGINAL (TIME VITOR)
 * Código original, extraído do documento do projeto.
 * E-mails e nomes reais foram trocados por exemplos: preencha
 * EMAILS_GESTORES, EMAILS_AGENTES e EMAILS_APRENDIZES antes de usar.
 ***************************************************************/

function salvarResultadoDoDia() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("BASE_CONSOLIDADA");
  const historico = ss.getSheetByName("HISTÓRICO");

  const dados = base.getRange(2, 1, 24, 12).getValues();

  const ontem = new Date();
  ontem.setDate(ontem.getDate() - 1);
  ontem.setHours(0, 0, 0, 0);

  const ultimaLinha = historico.getLastRow();

  if (ultimaLinha > 1) {
    const datasHistorico = historico.getRange(2, 1, ultimaLinha - 1, 1).getValues();

    const jaExiste = datasHistorico.some(linha => {
      const data = new Date(linha[0]);
      data.setHours(0, 0, 0, 0);
      return data.getTime() === ontem.getTime();
    });

    if (jaExiste) {
      mostrarAlerta("⚠ O resultado de ontem já foi salvo no HISTÓRICO.");
      return;
    }
  }

  const linhas = [];

  dados.forEach(linha => {
    if (linha[0] != "") {
      linhas.push([
        ontem,
        linha[0],
        linha[1],
        linha[2],
        linha[9],
        linha[10],
        linha[11],
        linha[4],
        linha[5],
        linha[7]
      ]);
    }
  });

  historico.getRange(
    historico.getLastRow() + 1,
    1,
    linhas.length,
    linhas[0].length
  ).setValues(linhas);

  mostrarAlerta("✅ Resultado salvo com sucesso!");
}

function importarCSVGenesysDoGmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawPerformance = ss.getSheetByName("RAW_PERFORMANCE");
  const rawOciosidade = ss.getSheetByName("RAW_OCIOSIDADE");

  const query = 'from:reporting-export-noreply@mypurecloud.com subject:"Relatório Genesys Cloud de Supervisor Vitor" has:attachment newer_than:2d';
  const threads = GmailApp.search(query);

  if (threads.length == 0) {
    mostrarAlerta("Nenhum relatório encontrado nos últimos dias.");
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const lastPerfId = props.getProperty("LAST_PERFORMANCE_ID");
  const lastOcioId = props.getProperty("LAST_OCIOSIDADE_ID");

  let perfIdNovo = null;
  let ocioIdNovo = null;
  let mudouAlgo = false;

  for (let t = 0; t < threads.length; t++) {
    const messages = threads[t].getMessages();
   
    for (let m = messages.length - 1; m >= 0; m--) {
      const message = messages[m];
      const msgId = message.getId();
      const anexos = message.getAttachments();
     
      for (let a = 0; a < anexos.length; a++) {
        const anexo = anexos[a];
        const nomeArquivo = anexo.getName().toLowerCase();

        if (nomeArquivo.indexOf("performance") > -1 && !perfIdNovo) {
          if (msgId === lastPerfId) {
            perfIdNovo = "JA_PROCESSADO";
          } else {
            importarCSV(anexo, rawPerformance);
            perfIdNovo = msgId;
            mudouAlgo = true;
          }
        }

        if (nomeArquivo.indexOf("ociosidade") > -1 && !ocioIdNovo) {
          if (msgId === lastOcioId) {
            ocioIdNovo = "JA_PROCESSADO";
          } else {
            importarCSV(anexo, rawOciosidade);
            ocioIdNovo = msgId;
            mudouAlgo = true;
          }
        }
      }
      if (perfIdNovo && ocioIdNovo) break;
    }
    if (perfIdNovo && ocioIdNovo) break;
  }

  if (!mudouAlgo) {
    Logger.log("Nenhum e-mail novo da Genesys chegou ainda. Aguardando próximo ciclo.");
    return;
  }

  if (perfIdNovo && perfIdNovo !== "JA_PROCESSADO") props.setProperty("LAST_PERFORMANCE_ID", perfIdNovo);
  if (ocioIdNovo && ocioIdNovo !== "JA_PROCESSADO") props.setProperty("LAST_OCIOSIDADE_ID", ocioIdNovo);

  const finalPerfId = props.getProperty("LAST_PERFORMANCE_ID");
  const finalOcioId = props.getProperty("LAST_OCIOSIDADE_ID");
  const sentPerfId = props.getProperty("SENT_PERFORMANCE_ID");
  const sentOcioId = props.getProperty("SENT_OCIOSIDADE_ID");

  if (finalPerfId !== sentPerfId && finalOcioId !== sentOcioId) {
    props.setProperty("SENT_PERFORMANCE_ID", finalPerfId);
    props.setProperty("SENT_OCIOSIDADE_ID", finalOcioId);

    mostrarAlerta(
      "Importação concluída! Ambos os relatórios do horário foram recebidos.\n\n" +
      "Disparando ranking e e-mails para o time..."
    );

    enviarEmailsERankingDoTime(false); // Falso = Envio real para a equipe

  } else {
    Logger.log("Planilha atualizada parcial. Aguardando o segundo relatório para disparar os e-mails.");
    mostrarAlerta("Aba atualizada com sucesso! Aguardando o par do relatório chegar para disparar o ranking.");
  }
}

function importarCSV(anexo, aba) {
  const csv = anexo.getDataAsString("UTF-8");
  const primeiraLinha = csv.split('\n')[0];
  const delimitador = primeiraLinha.indexOf(';') > -1 ? ';' : ',';
  const dados = Utilities.parseCsv(csv, delimitador);

  aba.clearContents();
  aba.getRange(1, 1, dados.length, dados[0].length).setValues(dados);
}

function dispararTesteParaVitor() {
  mostrarAlerta("⏳ Modo Teste Ativado!\n\nGerando a visão consolidada e a visão de CADA agente e aprendiz. Tudo será enviado APENAS para a sua caixa de entrada. A equipe não receberá nada.");
  enviarEmailsERankingDoTime(true);
}

function enviarEmailsERankingDoTime(somenteParaVitor = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const base = ss.getSheetByName("BASE_CONSOLIDADA");
  if (!base) return;

  // Lógica de Trava de Gestores
  let EMAILS_GESTORES = ["gestor@exemplo.com.br", "coordenacao@exemplo.com.br"];
  if (somenteParaVitor) {
    EMAILS_GESTORES = ["gestor@exemplo.com.br"];
  }

  const EMAILS_AGENTES = {
    // Preencha com "Nome exatamente como na Genesys": "e-mail"
    "Nome do Agente 1": "agente1@exemplo.com.br",
    "Nome do Agente 2": "agente2@exemplo.com.br",
  };

  const EMAILS_APRENDIZES = {
    // Preencha com "Nome exatamente como na Genesys": "e-mail"
    "Nome do Aprendiz 1": "aprendiz1@exemplo.com.br",
  };

  function normalizarNome(texto) {
    if (!texto) return "";
    return texto.toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim();
  }

  const EMAILS_NORMALIZADOS_AGENTES = {};
  for (let nome in EMAILS_AGENTES) { EMAILS_NORMALIZADOS_AGENTES[normalizarNome(nome)] = EMAILS_AGENTES[nome]; }

  const EMAILS_NORMALIZADOS_APRENDIZES = {};
  for (let nome in EMAILS_APRENDIZES) { EMAILS_NORMALIZADOS_APRENDIZES[normalizarNome(nome)] = EMAILS_APRENDIZES[nome]; }

  const ultimaLinha = base.getLastRow();
  if (ultimaLinha < 3) return;
  const dados = base.getRange(3, 1, ultimaLinha - 2, 10).getValues();

  // Arrays Separados
  let agentesAtivos = [];
  let agentesAusentes = [];
  let aprendizesAtivos = [];
  let aprendizesAusentes = [];

  let qtdEntregaram = 0;
  let qtdNaoEntregaram = 0;

  dados.forEach(linha => {
    if (!linha[0]) return;
    let nomeAgente = linha[0].toString().replace(/\n/g, "").trim();
    let producao = linha[1];      
    let metaOriginal = linha[2];  
    let metaAjustada = linha[8];  

    let nomeNorm = normalizarNome(nomeAgente);
    let ehAprendiz = !!EMAILS_NORMALIZADOS_APRENDIZES[nomeNorm];
    let emailOriginal = ehAprendiz ? EMAILS_NORMALIZADOS_APRENDIZES[nomeNorm] : EMAILS_NORMALIZADOS_AGENTES[nomeNorm];

    let isAusente = (producao === "" || producao === null || producao === undefined);
    let metaTratada = (metaAjustada !== "" && metaAjustada !== null) ? Number(metaAjustada) : ((metaOriginal !== "" && metaOriginal !== null) ? Number(metaOriginal) : 0);

    if (isAusente) {
      let objAusente = { nome: nomeAgente, producao: 0, meta: metaTratada, statusTabela: "Falta/Folga", corTabela: "#e8e8e8", emailOriginal: emailOriginal };
      if (ehAprendiz) aprendizesAusentes.push(objAusente);
      else agentesAusentes.push(objAusente);
    } else {
      let producaoTratada = Number(producao);
     
      let metaNuvens = metaTratada;
      let metaEstrelas = Math.round(metaTratada * 1.2);
     
      let restamNuvens = metaNuvens - producaoTratada;
      let restamEstrelas = metaEstrelas - producaoTratada;

      let fraseStatus = "";
      let statusTabela = "";
      let corTabela = "";

      if (restamNuvens > 0) {
        if (ehAprendiz) {
          fraseStatus = "para a Meta Nuvens (100%) faltam " + restamNuvens + " atendimentos e para a Meta Estrelas (120%) faltam " + restamEstrelas + " atendimentos. Acreditamos muito no seu potencial!";
          statusTabela = "Faltam " + restamNuvens + " Nuvens";
        } else {
          fraseStatus = "para a Meta Nuvens ☁️ (100%) faltam " + restamNuvens + " atendimentos e para a Meta Estrelas ⭐ (120%) faltam " + restamEstrelas + " atendimentos";
          statusTabela = "Faltam " + restamNuvens + " ☁️";
        }
        corTabela = "#ffebee";
        qtdNaoEntregaram++;
      } else if (restamEstrelas > 0) {
        if (ehAprendiz) {
          fraseStatus = "parabéns! Você BATEU a Meta Nuvens! Agora o seu desafio é a Meta Estrelas (120%), faltam apenas " + restamEstrelas + " atendimentos. Continue com esse foco maravilhoso!";
          statusTabela = "Nuvens OK! Faltam " + restamEstrelas + " Estrelas";
        } else {
          fraseStatus = "parabéns! Você BATEU a Meta Nuvens ☁️! Agora o seu desafio é a Meta Estrelas ⭐ (120%), faltam apenas " + restamEstrelas + " atendimentos";
          statusTabela = "Nuvens OK! Faltam " + restamEstrelas + " ⭐";
        }
        corTabela = "#e8f5e9";
        qtdEntregaram++;
      } else {
        if (ehAprendiz) {
          fraseStatus = "sensacional! Você superou todas as expectativas e BATEU as duas metas do dia: Nuvens e Estrelas! Você é um verdadeiro talento da nossa equipe!";
          statusTabela = "Dupla Meta Batida!";
        } else {
          fraseStatus = "sensacional! Você superou todas as expectativas e BATEU as duas metas do dia: Nuvens ☁️ e Estrelas ⭐! Você é gigante!";
          statusTabela = "Dupla Meta! ☁️⭐";
        }
        corTabela = "#e1f5fe";
        qtdEntregaram++;
      }

      let objAtivo = {
        nome: nomeAgente,
        producao: producaoTratada,
        meta: metaNuvens,
        metaEstrelas: metaEstrelas,
        fraseStatus: fraseStatus,
        statusTabela: statusTabela,
        corTabela: corTabela,
        emailOriginal: emailOriginal
      };

      if (ehAprendiz) aprendizesAtivos.push(objAtivo);
      else agentesAtivos.push(objAtivo);
    }
  });

  let totalOperando = qtdEntregaram + qtdNaoEntregaram;
  let pctEntregaPessoas = totalOperando > 0 ? (qtdEntregaram / totalOperando * 100) : 0;

  // Ordenação dos Rankings Separados
  agentesAtivos.sort((a, b) => b.producao - a.producao);
  aprendizesAtivos.sort((a, b) => b.producao - a.producao);

  let qtdAtivosAgentes = agentesAtivos.length;
  let corteTop5Agentes = qtdAtivosAgentes >= 5 ? agentesAtivos[4].producao : (qtdAtivosAgentes > 0 ? agentesAtivos[qtdAtivosAgentes - 1].producao : 0);

  let qtdAtivosAprendizes = aprendizesAtivos.length;
  let corteTop3Aprendizes = qtdAtivosAprendizes >= 3 ? aprendizesAtivos[2].producao : (qtdAtivosAprendizes > 0 ? aprendizesAtivos[qtdAtivosAprendizes - 1].producao : 0);

  let todosAtivos = [...agentesAtivos, ...aprendizesAtivos];
  let teamProd = todosAtivos.reduce((sum, a) => sum + a.producao, 0);
  let teamMeta = todosAtivos.reduce((sum, a) => sum + a.meta, 0);
  let teamPct = teamMeta > 0 ? (teamProd / teamMeta * 100) : 0;
  let teamCor = teamPct >= 100 ? "#2e7d32" : "#d32f2f";

  // --- MONTAGEM DO EMAIL DO GESTOR (CONSOLIDADO) ---
  let relatorioGestor = "<html><head><meta charset='UTF-8'></head><body style='font-family: Arial, sans-serif;'>" +
    "<h2>Resumo de Desempenho do Time - SAC Magalu</h2>" +
    "<div style='background-color: #f4f7f6; border-left: 6px solid #0056b3; padding: 15px; margin-bottom: 25px; border-radius: 4px; max-width: 600px;'>" +
    "  <h3 style='margin-top: 0; color: #0056b3;'>[Consolidado Geral da Equipe (Agentes + Aprendizes)]</h3>" +
    "  <p style='margin: 5px 0;'><b>Produção Total da Equipe:</b> " + teamProd + " atendimentos</p>" +
    "  <p style='margin: 5px 0;'><b>Meta Total da Equipe:</b> " + teamMeta + " atendimentos</p>" +
    "  <p style='margin: 5px 0;'><b>Atingimento Geral do Time:</b> <span style='color: " + teamCor + "; font-weight: bold; font-size: 16px;'>" + teamPct.toFixed(2) + "%</span></p>" +
    "  <p style='margin: 12px 0 5px 0; border-top: 1px dashed #cbd5e1; padding-top: 10px;'><b>Qtd. de Pessoas que Entregaram:</b> " + qtdEntregaram + " (SIM)</p>" +
    "  <p style='margin: 5px 0;'><b>Qtd. de Pessoas que NÃO Entregaram:</b> " + qtdNaoEntregaram + " (NÃO)</p>" +
    "  <p style='margin: 5px 0;'><b>Porcentagem de Entrega Operacional:</b> <span style='color: #0046b3; font-weight: bold; font-size: 15px;'>" + pctEntregaPessoas.toFixed(1) + "%</span></p>" +
    "</div>";

  // Tabela 1: Agentes
  relatorioGestor += "<h3 style='color: #333;'>Ranking Principal - Agentes</h3>" +
    "<table border='1' cellpadding='5' style='border-collapse: collapse; font-family: Arial, sans-serif; margin-bottom: 25px;'>" +
    "<tr style='background-color: #0056b3; color: white;'><th>Posição</th><th>Agente</th><th>Produção Atual</th><th>Meta Base (100%)</th><th>Status Metas</th></tr>";

  agentesAtivos.forEach((agente, indice) => {
    let posicao = indice + 1;
    relatorioGestor += "<tr style='background-color: " + agente.corTabela + ";'><td align='center'><b>" + posicao + "º</b></td><td>" + agente.nome + "</td><td>" + agente.producao + "</td><td>" + agente.meta + "</td><td><b>" + agente.statusTabela + "</b></td></tr>";
  });
  agentesAusentes.forEach(agente => {
    relatorioGestor += "<tr style='background-color: " + agente.corTabela + "; color: #777777;'><td align='center'>-</td><td>" + agente.nome + "</td><td align='center'>-</td><td>" + agente.meta + "</td><td>" + agente.statusTabela + "</td></tr>";
  });
  relatorioGestor += "</table>";

  // Tabela 2: Aprendizes
  relatorioGestor += "<h3 style='color: #333;'>Ranking Especial - Aprendizes</h3>" +
    "<table border='1' cellpadding='5' style='border-collapse: collapse; font-family: Arial, sans-serif; margin-bottom: 25px;'>" +
    "<tr style='background-color: #0086ff; color: white;'><th>Posição</th><th>Aprendiz</th><th>Produção Atual</th><th>Meta Base (100%)</th><th>Status Metas</th></tr>";

  aprendizesAtivos.forEach((aprendiz, indice) => {
    let posicao = indice + 1;
    relatorioGestor += "<tr style='background-color: " + aprendiz.corTabela + ";'><td align='center'><b>" + posicao + "º</b></td><td>" + aprendiz.nome + "</td><td>" + aprendiz.producao + "</td><td>" + aprendiz.meta + "</td><td><b>" + aprendiz.statusTabela + "</b></td></tr>";
  });
  aprendizesAusentes.forEach(aprendiz => {
    relatorioGestor += "<tr style='background-color: " + aprendiz.corTabela + "; color: #777777;'><td align='center'>-</td><td>" + aprendiz.nome + "</td><td align='center'>-</td><td>" + aprendiz.meta + "</td><td>" + aprendiz.statusTabela + "</td></tr>";
  });

  relatorioGestor += "</table>" +
    "<br><div style='font-family: Arial, sans-serif; font-weight: bold; color: #0046b3; font-size: 14px;'><span style='color: #0086ff;'>❤</span> TIME VITOR, o melhor time de todos! <span style='color: #0086ff;'>❤</span></div>" +
    "<p style='font-family: Arial, sans-serif; color: gray; font-size: 11px;'>Gerado automaticamente via Automação Planilha URA.</p></body></html>";

  let assuntoConsolidado = somenteParaVitor ? "[TESTE GESTOR] Visão Consolidada" : "[Geral Gestor] Acompanhamento de Metas e Rankings";
  EMAILS_GESTORES.forEach(gestor => { try { GmailApp.sendEmail(gestor, assuntoConsolidado, "", { htmlBody: relatorioGestor }); } catch(e){} });


  // --- DISPARO DOS E-MAILS INDIVIDUAIS: AGENTES ---
  agentesAtivos.forEach((agente, indice) => {
    let posicao = indice + 1;
    let rankingHtml = "";

    if (posicao === 1) {
      rankingHtml = "<span style='color: #b7791f; font-weight: bold;'>1º LUGAR! Você conquistou o topo do ranking do time!</span> Desempenho fantástico, continue liderando a fila!";
    } else if (posicao <= 5) {
      let dif = agentesAtivos[indice - 1].producao - agente.producao;
      rankingHtml = "⭐ <span style='color: #0056b3; font-weight: bold;'>Sensacional! Você faz parte do TOP 5 da equipe!</span><br>Sua posição atual é o <b>" + posicao + "º lugar</b>. Faltam apenas <b>" + dif + "</b> atendimentos para subir mais um nível!";
    } else {
      let dif5 = (corteTop5Agentes - agente.producao) + 1;
      rankingHtml = "Você está na <b>" + posicao + "ª posição</b> do ranking do time.<br>Faltam <b>" + dif5 + "</b> atendimentos para você entrar no seleto grupo do <b>TOP 5</b>! Você consegue, acelera!";
    }

    let emailFinalDeEnvio = somenteParaVitor ? "gestor@exemplo.com.br" : agente.emailOriginal;

    if (agente.emailOriginal) {
      let corpoAgente = "<html><head><meta charset='UTF-8'></head>" +
        "<body style='margin: 0; padding: 20px; background-color: #f4f7f6; font-family: Arial, sans-serif; color: #333333;'>" +
        "  <div style='max-width: 580px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.08); border: 1px solid #e2e8f0;'>" +
        "    " +
        "    <!-- Banner Superior Magalu -->" +
        "    <div style='background: linear-gradient(135deg, #0086ff 0%, #0046b3 100%); padding: 25px 20px; text-align: center; color: #ffffff;'>" +
        "      <h1 style='margin: 0; font-size: 22px; letter-spacing: 0.5px; font-weight: bold;'>[ FOCO NAS MÉTRICAS! ]</h1>" +
        "      <p style='margin: 4px 0 0 0; opacity: 0.85; font-size: 14px;'>Acompanhamento de Performance e Metas — SAC Magalu</p>" +
        "    </div>" +
        "    " +
        "    <!-- Corpo da Mensagem -->" +
        "    <div style='padding: 25px;'>" +
        "      <p style='font-size: 17px; margin-top: 0;'>Olá, <strong style='color: #0086ff;'>" + agente.nome + "</strong>!</p>" +
        "      <p style='font-size: 14px; line-height: 1.5; color: #4a5568;'>Aqui está o balanço dos seus indicadores atualizados agora mesmo. Veja os seus números e ajuste o seu ritmo para buscar os objetivos:</p>" +
        "      " +
        "      <!-- Card Box de Status de Entregas -->" +
        "      <div style='background-color: " + agente.corTabela + "; border-left: 6px solid #0086ff; padding: 15px; border-radius: 6px; margin: 20px 0; font-size: 14px; line-height: 1.5; font-weight: 500; color: #1a202c;'>" +
        "        \"" + agente.nome + ", " + agente.fraseStatus + ".\"" +
        "      </div>" +
        "      " +
        "      <!-- Card Box Interativo de Ranking -->" +
        "      <div style='background-color: #f0f7ff; border: 1px solid #b8daff; padding: 15px; border-radius: 8px; margin: 20px 0; text-align: center; font-size: 14px; line-height: 1.5; color: #2d3748;'>" +
                 rankingHtml +
        "      </div>" +
        "      " +
        "      <!-- Tabela de Métricas Técnicas -->" +
        "      <table style='width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #f8f9fa; border-radius: 6px; overflow: hidden;'>" +
        "        <tr style='background-color: #edf2f7; text-align: left; font-size: 12px; color: #4a5568; font-weight: bold;'>" +
        "          <th style='padding: 10px;'>Indicador de Desempenho</th>" +
        "          <th style='padding: 10px; text-align: center;'>Volumetria</th>" +
        "        </tr>" +
        "        <tr style='font-size: 13px; border-bottom: 1px solid #e2e8f0;'>" +
        "          <td style='padding: 10px; color: #4a5568;'>Produção Realizada até agora:</td>" +
        "          <td style='padding: 10px; text-align: center; font-weight: bold; color: #2d3748;'>" + agente.producao + "</td>" +
        "        </tr>" +
        "        <tr style='font-size: 13px; border-bottom: 1px solid #e2e8f0;'>" +
        "          <td style='padding: 10px; color: #0086ff; font-weight: bold;'>☁️ Alvo Meta Nuvens (100%):</td>" +
        "          <td style='padding: 10px; text-align: center; font-weight: bold; color: #0086ff;'>" + agente.meta + "</td>" +
        "        </tr>" +
        "        <tr style='font-size: 13px;'>" +
        "          <td style='padding: 10px; color: #b7791f; font-weight: bold;'>⭐ Alvo Meta Estrelas (120%):</td>" +
        "          <td style='padding: 10px; text-align: center; font-weight: bold; color: #b7791f;'>" + agente.metaEstrelas + "</td>" +
        "        </tr>" +
        "      </table>" +
        "      " +
        "      <!-- Painel do Indicador Consolidado do Time -->" +
        "      <div style='text-align: center; margin-top: 25px; padding-top: 15px; border-top: 1px solid #e2e8f0;'>" +
        "        <p style='color: #4a5568; font-size: 13px; margin: 0;'><b>Resultado da Equipe:</b> Nosso time já atingiu <span style='color: #2e7d32; font-weight: bold; font-size: 15px;'>" + teamPct.toFixed(1) + "%</span> da nossa meta consolidada do dia!</p>" +
        "        <p style='color: #4a5568; font-size: 13px; margin: 6px 0 0 0;'><b>Índice de Entrega:</b> " + pctEntregaPessoas.toFixed(0) + "% da equipe já entregaram a meta base de hoje (" + qtdEntregaram + " SIM / " + qtdNaoEntregaram + " NÃO)</p>" +
        "        <p style='color: #a0aec0; font-size: 12px; margin: 5px 0 0 0;'>Bora manter o foco operacional e fazer um grande fechamento! Boas chamadas!</p>" +
        "      </div>" +
        "    </div>" +
        "    " +
        "    <!-- Assinatura Exclusiva do Time -->" +
        "    <div style='background-color: #f7fafc; padding: 15px; text-align: center; border-top: 1px solid #e2e8f0;'>" +
        "      <p style='margin: 0; font-weight: bold; color: #0046b3; font-size: 14px; letter-spacing: 0.5px;'><span style='color: #0086ff;'>❤</span> TIME VITOR, o melhor time de todos! <span style='color: #0086ff;'>❤</span></p>" +
        "    </div>" +
        "    " +
        "    <!-- Rodapé Institucional -->" +
        "    <div style='background-color: #f7fafc; padding: 12px; text-align: center; border-top: 1px dashed #cbd5e1;'>" +
        "      <small style='color: #a0aec0; font-size: 11px;'>SAC — Magazine Luiza | Relatório automatizado via integração URA</small>" +
        "    </div>" +
        "  </div>" +
        "</body></html>";
     
      let assuntoAgente = somenteParaVitor ? "[TESTE VISÃO AGENTE] Como chegará para: " + agente.nome : "[SAC Magalu] Acompanhamento de Meta e Ranking Diário";
      try { GmailApp.sendEmail(emailFinalDeEnvio, assuntoAgente, "", { htmlBody: corpoAgente }); } catch(e){}
    }
  });


  // --- DISPARO DOS E-MAILS INDIVIDUAIS: APRENDIZES ---
  aprendizesAtivos.forEach((aprendiz, indice) => {
    let posicao = indice + 1;
    let rankingHtml = "";

    if (posicao === 1) {
      rankingHtml = "<span style='color: #FF0080; font-weight: bold; font-size: 16px;'>1º LUGAR! Você é o grande destaque dos aprendizes hoje!</span><br>Que orgulho, continue dando o seu melhor!";
    } else if (posicao <= 3) {
      let dif = aprendizesAtivos[indice - 1].producao - aprendiz.producao;
      rankingHtml = "<span style='color: #0086ff; font-weight: bold; font-size: 16px;'>Incrível! Você está no TOP 3 dos aprendizes!</span><br>Sua posição atual é o <b>" + posicao + "º lugar</b>. Faltam apenas <b>" + dif + "</b> atendimentos para subir de nível!";
    } else {
      let dif3 = (corteTop3Aprendizes - aprendiz.producao) + 1;
      rankingHtml = "Você está na <b>" + posicao + "ª posição</b> do ranking dos aprendizes.<br>Faltam <b>" + dif3 + "</b> atendimentos para você entrar no <b>TOP 3</b>! Acelera que dá!";
    }

    let emailFinalDeEnvio = somenteParaVitor ? "gestor@exemplo.com.br" : aprendiz.emailOriginal;

    if (aprendiz.emailOriginal) {
      // E-mail Colorido Magalu e SEM EMOJIS!
      let corpoAprendiz = "<html><head><meta charset='UTF-8'></head>" +
        "<body style='margin: 0; padding: 20px; background-color: #e6f2ff; font-family: Arial, sans-serif; color: #333333;'>" +
        "  <div style='max-width: 580px; margin: 0 auto; background-color: #ffffff; border-radius: 15px; overflow: hidden; box-shadow: 0 8px 16px rgba(0,0,0,0.1); border: 2px solid #FFD700;'>" +
        "    " +
        "    <!-- Banner Superior Colorido Magalu -->" +
        "    <div style='background: linear-gradient(90deg, #0086ff 0%, #FF0080 100%); padding: 30px 20px; text-align: center; color: #ffffff;'>" +
        "      <h1 style='margin: 0; font-size: 26px; font-weight: bold; text-shadow: 1px 1px 2px rgba(0,0,0,0.2);'>[ HORA DO SHOW! ]</h1>" +
        "      <p style='margin: 8px 0 0 0; font-size: 16px; font-weight: bold;'>Seu Acompanhamento de Performance</p>" +
        "    </div>" +
        "    " +
        "    <!-- Corpo da Mensagem Aprendiz -->" +
        "    <div style='padding: 30px;'>" +
        "      <p style='font-size: 18px; margin-top: 0; color: #0086ff; font-weight: bold;'>E aí, " + aprendiz.nome + "! Tudo bem?</p>" +
        "      <p style='font-size: 15px; line-height: 1.6; color: #555555;'>Olha só como estão os seus números! Você está indo super bem, continue brilhando e aprendendo muito com a gente!</p>" +
        "      " +
        "      <!-- Card Box de Status -->" +
        "      <div style='background-color: " + aprendiz.corTabela + "; border-left: 8px solid #FFD700; padding: 18px; border-radius: 8px; margin: 25px 0; font-size: 15px; line-height: 1.6; font-weight: bold; color: #333; box-shadow: 0 2px 5px rgba(0,0,0,0.05);'>" +
        "        \"" + aprendiz.nome + ", " + aprendiz.fraseStatus + "\"" +
        "      </div>" +
        "      " +
        "      <!-- Card Box Interativo de Ranking -->" +
        "      <div style='background-color: #e6f7ff; border: 2px dashed #0086ff; padding: 18px; border-radius: 12px; margin: 25px 0; text-align: center; font-size: 15px; line-height: 1.6; color: #333;'>" +
                 rankingHtml +
        "      </div>" +
        "      " +
        "      <!-- Tabela de Métricas -->" +
        "      <table style='width: 100%; border-collapse: separate; border-spacing: 0; margin: 25px 0; border-radius: 8px; overflow: hidden; border: 1px solid #e2e8f0;'>" +
        "        <tr style='background-color: #0086ff; color: white; text-align: left; font-size: 14px;'>" +
        "          <th style='padding: 12px;'>Seu Desempenho</th>" +
        "          <th style='padding: 12px; text-align: center;'>Volumetria</th>" +
        "        </tr>" +
        "        <tr style='font-size: 14px; background-color: #f8f9fa;'>" +
        "          <td style='padding: 12px; border-bottom: 1px solid #e2e8f0;'>Produção Realizada:</td>" +
        "          <td style='padding: 12px; text-align: center; font-weight: bold; border-bottom: 1px solid #e2e8f0; color: #333;'>" + aprendiz.producao + "</td>" +
        "        </tr>" +
        "        <tr style='font-size: 14px; background-color: #ffffff;'>" +
        "          <td style='padding: 12px; color: #0086ff; font-weight: bold; border-bottom: 1px solid #e2e8f0;'>Alvo Meta Nuvens (100%):</td>" +
        "          <td style='padding: 12px; text-align: center; font-weight: bold; color: #0086ff; border-bottom: 1px solid #e2e8f0;'>" + aprendiz.meta + "</td>" +
        "        </tr>" +
        "        <tr style='font-size: 14px; background-color: #f8f9fa;'>" +
        "          <td style='padding: 12px; color: #FF0080; font-weight: bold;'>Alvo Meta Estrelas (120%):</td>" +
        "          <td style='padding: 12px; text-align: center; font-weight: bold; color: #FF0080;'>" + aprendiz.metaEstrelas + "</td>" +
        "        </tr>" +
        "      </table>" +
        "      " +
        "      <!-- Painel do Indicador Consolidado -->" +
        "      <div style='text-align: center; margin-top: 30px; padding-top: 20px; border-top: 2px solid #f0f0f0;'>" +
        "        <p style='color: #666; font-size: 14px; margin: 0;'><b>Resultado da Equipe:</b> Nosso time já atingiu <span style='color: #00B25A; font-weight: bold; font-size: 16px;'>" + teamPct.toFixed(1) + "%</span> da nossa meta consolidada do dia!</p>" +
        "        <p style='color: #888; font-size: 13px; margin: 10px 0 0 0;'>Vamos juntos fazer um dia maravilhoso! Boa sorte nos aprendizados!</p>" +
        "      </div>" +
        "    </div>" +
        "    " +
        "    <!-- Assinatura Exclusiva Aprendizes (Sem Emojis) -->" +
        "    <div style='background-color: #FFD700; padding: 18px; text-align: center;'>" +
        "      <p style='margin: 0; font-weight: bold; color: #0046b3; font-size: 15px; text-transform: uppercase;'> TIME VITOR, O MELHOR TIME DE TODOS! </p>" +
        "    </div>" +
        "    " +
        "    <!-- Rodapé -->" +
        "    <div style='background-color: #f0f8ff; padding: 12px; text-align: center;'>" +
        "      <small style='color: #999; font-size: 11px;'>SAC — Magazine Luiza | Relatório automatizado via integração URA</small>" +
        "    </div>" +
        "  </div>" +
        "</body></html>";

      let assuntoAprendiz = somenteParaVitor ? "[TESTE VISÃO APRENDIZ] Como chegará para: " + aprendiz.nome : "[SAC Magalu] Acompanhamento Especial de Metas";
      try { GmailApp.sendEmail(emailFinalDeEnvio, assuntoAprendiz, "", { htmlBody: corpoAprendiz }); } catch(e){}
    }
  });
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("📊 Gestão Time Vitor")
    .addItem("🔄 Importar CSVs do Gmail", "importarCSVGenesysDoGmail")
    .addItem("📥 Salvar Resultado do Dia", "salvarResultadoDoDia")
    .addSeparator()
    .addItem("📩 Enviar Teste (Só para Vitor)", "dispararTesteParaVitor")
    .addToUi();
}

function mostrarAlerta(mensagem) {
  try { SpreadsheetApp.getUi().alert(mensagem); } catch (e) { Logger.log("Plano de fundo: " + mensagem.replace(/\n/g, " ")); }
}

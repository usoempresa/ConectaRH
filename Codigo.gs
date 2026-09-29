// ============================================================
// CONECTARH
// Sistema de comunicação e RH para 28 lojas
// Google Apps Script + Google Sheets
// ============================================================

const CONFIG = {
  NOME_SISTEMA: "ConectaRH",
  TEMPO_SESSAO: 21600, // 6 horas
  TOTAL_LOJAS: 28
};

// ============================================================
// ABAS
// ============================================================

const ABAS = {
  USUARIOS: "Usuarios",
  LOJAS: "Lojas",
  COLABORADORES: "Colaboradores",
  AVISOS: "Avisos",
  EVENTOS: "Eventos",
  DOCUMENTOS: "Documentos",
  FEEDBACKS: "Feedbacks",
  RECONHECIMENTOS: "Reconhecimentos",
  PESQUISAS: "Pesquisas",
  RESPOSTAS_PESQUISA: "RespostasPesquisa"
};

// ============================================================
// ABRIR SISTEMA
// ============================================================

function doGet() {

  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle(CONFIG.NOME_SISTEMA)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// PLANILHA
// ============================================================

function obterPlanilha() {

  return SpreadsheetApp.getActiveSpreadsheet();

}


// ============================================================
// CONFIGURAÇÃO INICIAL
// ============================================================

function configurarConectaRH() {

  const ss = obterPlanilha();

  criarAba(
    ss,
    ABAS.USUARIOS,
    [
      "ID",
      "CPF",
      "Senha",
      "Nome",
      "Perfil",
      "LojaID",
      "Ativo",
      "PrimeiroAcesso",
      "CriadoEm"
    ]
  );

  criarAba(
    ss,
    ABAS.LOJAS,
    [
      "ID",
      "Nome",
      "Codigo",
      "Cidade",
      "Ativa"
    ]
  );

  criarAba(
    ss,
    ABAS.COLABORADORES,
    [
      "ID",
      "CPF",
      "Nome",
      "DataNascimento",
      "Cargo",
      "Setor",
      "LojaID",
      "Telefone",
      "Email",
      "Foto",
      "Ativo",
      "CriadoEm"
    ]
  );

  criarAba(
    ss,
    ABAS.AVISOS,
    [
      "ID",
      "Titulo",
      "Texto",
      "LojaID",
      "Autor",
      "Data",
      "Importante",
      "Ativo"
    ]
  );

  criarAba(
    ss,
    ABAS.EVENTOS,
    [
      "ID",
      "Titulo",
      "Descricao",
      "Data",
      "Hora",
      "Local",
      "LojaID",
      "Autor",
      "Ativo"
    ]
  );

  criarAba(
    ss,
    ABAS.DOCUMENTOS,
    [
      "ID",
      "CPF",
      "Nome",
      "LojaID",
      "Tipo",
      "Descricao",
      "DataSolicitacao",
      "Status",
      "RespostaRH",
      "LinkDocumento",
      "DataResposta",
      "RespondidoPor"
    ]
  );

  criarAba(
    ss,
    ABAS.FEEDBACKS,
    [
      "ID",
      "CPF",
      "Nome",
      "LojaID",
      "Tipo",
      "Mensagem",
      "Data",
      "Lido",
      "Publicado"
    ]
  );

  criarAba(
    ss,
    ABAS.RECONHECIMENTOS,
    [
      "ID",
      "DeCPF",
      "DeNome",
      "ParaCPF",
      "ParaNome",
      "LojaID",
      "Mensagem",
      "Data",
      "Ativo"
    ]
  );

  criarAba(
    ss,
    ABAS.PESQUISAS,
    [
      "ID",
      "Categoria",
      "Pergunta",
      "Ativa",
      "Data",
      "LojaID",
      "Autor"
    ]
  );

  criarAba(
    ss,
    ABAS.RESPOSTAS_PESQUISA,
    [
      "ID",
      "PesquisaID",
      "CPF",
      "Nome",
      "LojaID",
      "Nota",
      "Data"
    ]
  );

  criarLojasPadrao();

  criarAdminInicial();

  criarPesquisasPadrao();

  return {
    sucesso: true,
    mensagem: "ConectaRH configurado com sucesso."
  };
}


// ============================================================
// CRIAR ABA
// ============================================================

function criarAba(ss, nome, cabecalhos) {

  let aba = ss.getSheetByName(nome);

  if (!aba) {
    aba = ss.insertSheet(nome);
  }

  if (aba.getLastRow() === 0) {

    aba
      .getRange(1, 1, 1, cabecalhos.length)
      .setValues([cabecalhos]);

    aba.getRange(1, 1, 1, cabecalhos.length)
      .setFontWeight("bold");

    aba.setFrozenRows(1);
  }
}


// ============================================================
// CRIAR 28 LOJAS
// ============================================================

function criarLojasPadrao() {

  const aba = obterAba(ABAS.LOJAS);

  if (aba.getLastRow() > 1) {
    return;
  }

  const dados = [];

  for (let i = 1; i <= CONFIG.TOTAL_LOJAS; i++) {

    dados.push([
      gerarID(),
      "Loja " + i,
      String(i).padStart(2, "0"),
      "",
      "SIM"
    ]);

  }

  aba
    .getRange(2, 1, dados.length, dados[0].length)
    .setValues(dados);
}


// ============================================================
// ADMIN INICIAL
// ============================================================

function criarAdminInicial() {

  const aba = obterAba(ABAS.USUARIOS);

  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][4]).toUpperCase() === "ADMIN"
    ) {
      return;
    }

  }

  const cpf = "00000000000";

  aba.appendRow([
    gerarID(),
    cpf,
    gerarHash(cpf),
    "Administrador ConectaRH",
    "ADMIN",
    "",
    "SIM",
    "SIM",
    new Date()
  ]);
}


// ============================================================
// CRIAR PESQUISAS PADRÃO
// ============================================================

function criarPesquisasPadrao() {

  const aba = obterAba(ABAS.PESQUISAS);

  if (aba.getLastRow() > 1) {
    return;
  }

  const categorias = [
    "Liderança",
    "Ambiente de trabalho",
    "Comunicação",
    "Benefícios",
    "Crescimento pessoal",
    "Equilíbrio vida-trabalho"
  ];

  const perguntas = [
    "Como você avalia a liderança da empresa?",
    "Como você avalia o ambiente de trabalho?",
    "Como você avalia a comunicação interna?",
    "Como você avalia os benefícios oferecidos?",
    "Como você avalia as oportunidades de crescimento pessoal?",
    "Como você avalia o equilíbrio entre vida pessoal e trabalho?"
  ];

  const dados = [];

  for (let i = 0; i < categorias.length; i++) {

    dados.push([
      gerarID(),
      categorias[i],
      perguntas[i],
      "SIM",
      new Date(),
      "",
      "Sistema"
    ]);

  }

  aba
    .getRange(2, 1, dados.length, dados[0].length)
    .setValues(dados);
}


// ============================================================
// UTILITÁRIOS
// ============================================================

function obterAba(nome) {

  const aba = obterPlanilha().getSheetByName(nome);

  if (!aba) {
    throw new Error(
      "A aba '" + nome + "' não existe. Execute configurarConectaRH()."
    );
  }

  return aba;
}


function gerarID() {

  return Utilities.getUuid();
}


function normalizarCPF(cpf) {

  return String(cpf || "")
    .replace(/\D/g, "")
    .padStart(11, "0");
}


function gerarHash(texto) {

  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(texto)
  );

  return bytes
    .map(function(byte) {

      const v = byte < 0 ? byte + 256 : byte;

      return ("0" + v.toString(16)).slice(-2);

    })
    .join("");
}


function agora() {

  return new Date();
}


function formatarData(data) {

  if (!data) {
    return "";
  }

  try {

    return Utilities.formatDate(
      new Date(data),
      Session.getScriptTimeZone() || "America/Sao_Paulo",
      "dd/MM/yyyy"
    );

  } catch (e) {

    return "";

  }
}


function formatarDataHora(data) {

  if (!data) {
    return "";
  }

  try {

    return Utilities.formatDate(
      new Date(data),
      Session.getScriptTimeZone() || "America/Sao_Paulo",
      "dd/MM/yyyy HH:mm"
    );

  } catch (e) {

    return "";

  }
}


// ============================================================
// SESSÃO
// ============================================================

function criarSessao(usuario) {

  const token = Utilities.getUuid();

  CacheService
    .getScriptCache()
    .put(
      "CONECTARH_" + token,
      JSON.stringify(usuario),
      CONFIG.TEMPO_SESSAO
    );

  return token;
}


function obterSessao(token) {

  if (!token) {
    throw new Error("Sessão não informada.");
  }

  const dados = CacheService
    .getScriptCache()
    .get("CONECTARH_" + token);

  if (!dados) {
    throw new Error("Sessão expirada. Faça login novamente.");
  }

  return JSON.parse(dados);
}


function encerrarSessao(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove("CONECTARH_" + token);

  }

  return true;
}


// ============================================================
// LOGIN
// ============================================================

function login(cpf, senha) {

  cpf = normalizarCPF(cpf);

  senha = String(senha || "");

  if (!cpf || cpf.length !== 11) {

    return {
      sucesso: false,
      mensagem: "Informe um CPF válido."
    };

  }

  const aba = obterAba(ABAS.USUARIOS);

  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    const cpfBanco = normalizarCPF(linha[1]);

    if (cpfBanco === cpf) {

      if (String(linha[6]).toUpperCase() !== "SIM") {

        return {
          sucesso: false,
          mensagem: "Usuário desativado."
        };

      }

      const senhaHash = gerarHash(senha);

      if (senhaHash !== String(linha[2])) {

        return {
          sucesso: false,
          mensagem: "CPF ou senha incorretos."
        };

      }

      const usuario = {

        id: linha[0],
        cpf: cpf,
        nome: linha[3],
        perfil: String(linha[4]).toUpperCase(),
        lojaID: linha[5],
        primeiroAcesso:
          String(linha[7]).toUpperCase() === "SIM"

      };

      const token = criarSessao(usuario);

      return {
        sucesso: true,
        token: token,
        usuario: usuario
      };

    }

  }

  return {
    sucesso: false,
    mensagem: "CPF ou senha incorretos."
  };
}


// ============================================================
// ALTERAR SENHA
// ============================================================

function alterarSenha(token, senhaAtual, novaSenha) {

  const usuario = obterSessao(token);

  if (!novaSenha || String(novaSenha).length < 6) {

    throw new Error(
      "A nova senha precisa ter pelo menos 6 caracteres."
    );

  }

  const aba = obterAba(ABAS.USUARIOS);

  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    if (normalizarCPF(dados[i][1]) === usuario.cpf) {

      if (
        gerarHash(senhaAtual) !==
        String(dados[i][2])
      ) {

        throw new Error("Senha atual incorreta.");

      }

      aba.getRange(i + 1, 3)
        .setValue(gerarHash(novaSenha));

      aba.getRange(i + 1, 8)
        .setValue("NAO");

      usuario.primeiroAcesso = false;

      CacheService
        .getScriptCache()
        .put(
          "CONECTARH_" + token,
          JSON.stringify(usuario),
          CONFIG.TEMPO_SESSAO
        );

      return {
        sucesso: true,
        mensagem: "Senha alterada com sucesso."
      };

    }

  }

  throw new Error("Usuário não encontrado.");
}


// ============================================================
// DASHBOARD
// ============================================================

function carregarDashboard(token) {

  const usuario = obterSessao(token);

  const resultado = {

    usuario: usuario,

    avisos: listarAvisos(token).slice(0, 5),

    eventos: listarEventos(token).slice(0, 5),

    aniversariantes:
      listarAniversariantes(token),

    documentosPendentes: 0,

    feedbacksPendentes: 0,

    colaboradores: 0

  };


  if (
    usuario.perfil === "RH" ||
    usuario.perfil === "ADMIN"
  ) {

    resultado.documentosPendentes =
      contarDocumentosPendentes(token);

    resultado.feedbacksPendentes =
      contarFeedbacksPendentes(token);

    resultado.colaboradores =
      listarColaboradores(token).length;

  }


  return resultado;
}


// ============================================================
// VERIFICAÇÃO DE LOJA
// ============================================================

function usuarioPodeVerLoja(usuario, lojaID) {

  if (usuario.perfil === "ADMIN") {
    return true;
  }

  // Aviso geral para todas as lojas
  if (!lojaID || String(lojaID).toUpperCase() === "TODAS") {
    return true;
  }

  return String(usuario.lojaID) === String(lojaID);
}


// ============================================================
// AVISOS
// ============================================================

function listarAvisos(token) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.AVISOS);

  const dados = aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (String(linha[7]).toUpperCase() !== "SIM") {
      continue;
    }

    if (!usuarioPodeVerLoja(usuario, linha[3])) {
      continue;
    }

    resultado.push({

      id: linha[0],
      titulo: linha[1],
      texto: linha[2],
      lojaID: linha[3],
      autor: linha[4],
      data: formatarDataHora(linha[5]),
      importante:
        String(linha[6]).toUpperCase() === "SIM"

    });

  }

  resultado.reverse();

  return resultado;
}


function criarAviso(token, dados) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Somente o RH ou Administrador pode publicar avisos."
    );

  }

  let lojaID;

  if (usuario.perfil === "ADMIN") {
    lojaID = dados.lojaID || "TODAS";
  } else {
    lojaID = dados.lojaID || usuario.lojaID;
  }

  if (!lojaID) {
    throw new Error("Informe a loja.");
  }

  const aba = obterAba(ABAS.AVISOS);

  aba.appendRow([
    gerarID(),
    dados.titulo,
    dados.texto,
    lojaID,
    usuario.nome,
    agora(),
    dados.importante ? "SIM" : "NAO",
    "SIM"
  ]);

  return {
    sucesso: true,
    mensagem: "Aviso publicado."
  };
}


// ============================================================
// EVENTOS
// ============================================================

function listarEventos(token) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.EVENTOS);

  const dados = aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (String(linha[8]).toUpperCase() !== "SIM") {
      continue;
    }

    if (!usuarioPodeVerLoja(usuario, linha[6])) {
      continue;
    }

    resultado.push({

      id: linha[0],
      titulo: linha[1],
      descricao: linha[2],
      data: formatarData(linha[3]),
      hora: linha[4],
      local: linha[5],
      lojaID: linha[6],
      autor: linha[7]

    });

  }

  resultado.reverse();

  return resultado;
}


function criarEvento(token, dados) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Somente o RH ou Administrador pode criar eventos."
    );

  }

  const lojaID =
    usuario.perfil === "ADMIN"
      ? dados.lojaID
      : usuario.lojaID;

  const aba = obterAba(ABAS.EVENTOS);

  aba.appendRow([
    gerarID(),
    dados.titulo,
    dados.descricao || "",
    dados.data,
    dados.hora || "",
    dados.local || "",
    lojaID,
    usuario.nome,
    "SIM"
  ]);

  return {
    sucesso: true,
    mensagem: "Evento criado."
  };
}


// ============================================================
// ANIVERSARIANTES
// ============================================================

function listarAniversariantes(token) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.COLABORADORES);

  const dados = aba.getDataRange().getValues();

  const mesAtual =
    new Date().getMonth();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (String(linha[10]).toUpperCase() !== "SIM") {
      continue;
    }

    if (!usuarioPodeVerLoja(usuario, linha[6])) {
      continue;
    }

    const nascimento = linha[3];

    if (!nascimento) {
      continue;
    }

    const dataNascimento = new Date(nascimento);

    if (
      dataNascimento.getMonth() === mesAtual
    ) {

      resultado.push({

        id: linha[0],
        nome: linha[2],
        data:
          Utilities.formatDate(
            dataNascimento,
            Session.getScriptTimeZone() ||
              "America/Sao_Paulo",
            "dd/MM"
          ),
        cargo: linha[4],
        setor: linha[5],
        foto: linha[9] || ""

      });

    }

  }

  resultado.sort(function(a, b) {

    return a.data.localeCompare(
      b.data
    );

  });

  return resultado;
}


// ============================================================
// COLABORADORES
// ============================================================

function listarColaboradores(token) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Acesso permitido somente ao RH."
    );

  }

  const aba =
    obterAba(ABAS.COLABORADORES);

  const dados =
    aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (!usuarioPodeVerLoja(usuario, linha[6])) {
      continue;
    }

    resultado.push({

      id: linha[0],
      cpf: linha[1],
      nome: linha[2],
      nascimento: formatarData(linha[3]),
      cargo: linha[4],
      setor: linha[5],
      lojaID: linha[6],
      telefone: linha[7],
      email: linha[8],
      foto: linha[9],
      ativo:
        String(linha[10]).toUpperCase() === "SIM"

    });

  }

  return resultado;
}


// ============================================================
// DOCUMENTOS
// ============================================================

function listarDocumentos(token) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.DOCUMENTOS);

  const dados = aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (usuario.perfil === "COLABORADOR") {

      if (normalizarCPF(linha[1]) !== usuario.cpf) {
        continue;
      }

    } else if (usuario.perfil !== "ADMIN" && usuario.perfil !== "RH") {

      continue;

    }

    resultado.push({

      id: linha[0],
      cpf: linha[1],
      nome: linha[2],
      lojaID: linha[3],
      tipo: linha[4],
      descricao: linha[5],
      dataSolicitacao: formatarData(linha[6]),
      status: linha[7],
      respostaRH: linha[8],
      linkDocumento: linha[9],
      dataResposta: formatarData(linha[10]),
      respondidoPor: linha[11]

    });

  }

  return resultado;
}


function solicitarDocumento(token, dados) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.DOCUMENTOS);

  aba.appendRow([
    gerarID(),
    usuario.cpf,
    usuario.nome,
    usuario.lojaID,
    dados.tipo,
    dados.descricao,
    agora(),
    "Pendente",
    "",
    "",
    "",
    ""
  ]);

  return {
    sucesso: true,
    mensagem: "Solicitação enviada com sucesso."
  };
}


function contarDocumentosPendentes(token) {

  const aba = obterAba(ABAS.DOCUMENTOS);

  const dados = aba.getDataRange().getValues();

  let contador = 0;

  for (let i = 1; i < dados.length; i++) {

    if (String(dados[i][7]).toUpperCase() === "PENDENTE") {
      contador++;
    }

  }

  return contador;
}


// ============================================================
// FEEDBACKS
// ============================================================

function listarFeedbacks(token) {

  const usuario = obterSessao(token);

  // Apenas RH e ADMIN podem ver feedbacks
  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Acesso permitido somente ao RH."
    );

  }

  const aba = obterAba(ABAS.FEEDBACKS);

  const dados = aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (!usuarioPodeVerLoja(usuario, linha[3])) {
      continue;
    }

    resultado.push({

      id: linha[0],
      cpf: linha[1],
      nome: linha[2],
      lojaID: linha[3],
      tipo: linha[4],
      mensagem: linha[5],
      data: formatarDataHora(linha[6]),
      lido: String(linha[7]).toUpperCase() === "SIM",
      publicado: String(linha[8]).toUpperCase() === "SIM"

    });

  }

  resultado.reverse();

  return resultado;
}


function enviarFeedback(token, dados) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.FEEDBACKS);

  aba.appendRow([
    gerarID(),
    usuario.cpf,
    usuario.nome,
    usuario.lojaID,
    dados.tipo,
    dados.mensagem,
    agora(),
    "NAO",
    "NAO"
  ]);

  return {
    sucesso: true,
    mensagem: "Feedback enviado com sucesso."
  };
}


function contarFeedbacksPendentes(token) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    return 0;

  }

  const aba = obterAba(ABAS.FEEDBACKS);

  const dados = aba.getDataRange().getValues();

  let contador = 0;

  for (let i = 1; i < dados.length; i++) {

    if (
      !usuarioPodeVerLoja(usuario, dados[i][3]) ||
      String(dados[i][7]).toUpperCase() === "SIM"
    ) {
      continue;
    }

    contador++;

  }

  return contador;
}


function marcarFeedbackLido(token, feedbackID) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Acesso permitido somente ao RH."
    );

  }

  const aba = obterAba(ABAS.FEEDBACKS);

  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    if (dados[i][0] === feedbackID) {

      aba.getRange(i + 1, 8).setValue("SIM");

      return {
        sucesso: true,
        mensagem: "Feedback marcado como lido."
      };

    }

  }

  throw new Error("Feedback não encontrado.");
}


function publicarFeedback(token, feedbackID) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Acesso permitido somente ao RH."
    );

  }

  const aba = obterAba(ABAS.FEEDBACKS);

  const dados = aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    if (dados[i][0] === feedbackID) {

      aba.getRange(i + 1, 9).setValue("SIM");

      aba.getRange(i + 1, 8).setValue("SIM");

      return {
        sucesso: true,
        mensagem: "Feedback publicado com sucesso."
      };

    }

  }

  throw new Error("Feedback não encontrado.");
}


// ============================================================
// RECONHECIMENTOS
// ============================================================

function listarReconhecimentos(token) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.RECONHECIMENTOS);

  const dados = aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (String(linha[8]).toUpperCase() !== "SIM") {
      continue;
    }

    if (!usuarioPodeVerLoja(usuario, linha[5])) {
      continue;
    }

    resultado.push({

      id: linha[0],
      de: linha[2],
      para: linha[4],
      mensagem: linha[6],
      data: formatarData(linha[7])

    });

  }

  resultado.reverse();

  return resultado;
}


function criarReconhecimento(token, dados) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.RECONHECIMENTOS);

  aba.appendRow([
    gerarID(),
    usuario.cpf,
    usuario.nome,
    dados.paraID || "",
    dados.paraNome || "",
    usuario.lojaID,
    dados.mensagem,
    agora(),
    "SIM"
  ]);

  return {
    sucesso: true,
    mensagem: "Reconhecimento enviado com sucesso."
  };
}


// ============================================================
// PESQUISAS
// ============================================================

function listarPesquisas(token) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.PESQUISAS);

  const dados = aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (String(linha[3]).toUpperCase() !== "SIM") {
      continue;
    }

    if (!usuarioPodeVerLoja(usuario, linha[4])) {
      continue;
    }

    resultado.push({

      id: linha[0],
      categoria: linha[1],
      pergunta: linha[2],
      data: formatarData(linha[4]),
      autor: linha[6]

    });

  }

  return resultado;
}


function responderPesquisa(token, dados) {

  const usuario = obterSessao(token);

  const aba = obterAba(ABAS.RESPOSTAS_PESQUISA);

  const todasRespostas = aba.getDataRange().getValues();

  // Verificar se o usuário já respondeu esta pesquisa
  for (let i = 1; i < todasRespostas.length; i++) {

    if (
      todasRespostas[i][1] === dados.pesquisaID &&
      normalizarCPF(todasRespostas[i][2]) === usuario.cpf
    ) {

      return {
        sucesso: false,
        mensagem: "Você já respondeu esta pesquisa."
      };

    }

  }

  aba.appendRow([
    gerarID(),
    dados.pesquisaID,
    usuario.cpf,
    usuario.nome,
    usuario.lojaID,
    dados.nota,
    agora()
  ]);

  return {
    sucesso: true,
    mensagem: "Resposta registrada com sucesso."
  };
}


function obterRespostasPesquisa(token, pesquisaID) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Acesso permitido somente ao RH."
    );

  }

  const aba = obterAba(ABAS.RESPOSTAS_PESQUISA);

  const dados = aba.getDataRange().getValues();

  const resultado = [];

  let somaNotas = 0;
  let totalRespostas = 0;

  for (let i = 1; i < dados.length; i++) {

    if (dados[i][1] === pesquisaID) {

      resultado.push({

        nome: dados[i][3],
        nota: dados[i][5],
        data: formatarData(dados[i][6])

      });

      somaNotas += parseInt(dados[i][5]) || 0;
      totalRespostas++;

    }

  }

  return {
    respostas: resultado,
    mediaNotas: totalRespostas > 0 ? (somaNotas / totalRespostas).toFixed(1) : 0,
    totalRespostas: totalRespostas
  };
}

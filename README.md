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
      "Lido"
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
      "Pergunta",
      "Opcao1",
      "Opcao2",
      "Opcao3",
      "Opcao4",
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
      "Resposta",
      "Data"
    ]
  );

  criarLojasPadrao();

  criarAdminInicial();

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
      "CONCERTARH_" + token,
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
    .get("CONCERTARH_" + token);

  if (!dados) {
    throw new Error("Sessão expirada. Faça login novamente.");
  }

  return JSON.parse(dados);
}


function encerrarSessao(token) {

  if (token) {

    CacheService
      .getScriptCache()
      .remove("CONCERTARH_" + token);

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
          "CONCERTARH_" + token,
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
      contarFeedbacks(token);

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

  const lojaID =
    usuario.perfil === "ADMIN"
      ? dados.lojaID
      : usuario.lojaID;

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
// CADASTRAR COLABORADOR
// ============================================================

function cadastrarColaborador(token, dados) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Somente o RH pode cadastrar colaboradores."
    );

  }

  const cpf = normalizarCPF(dados.cpf);

  if (cpf.length !== 11) {
    throw new Error("CPF inválido.");
  }

  const lojaID =
    usuario.perfil === "ADMIN"
      ? dados.lojaID
      : usuario.lojaID;

  if (!lojaID) {
    throw new Error("Loja não informada.");
  }

  const aba =
    obterAba(ABAS.COLABORADORES);

  const dadosExistentes =
    aba.getDataRange().getValues();

  for (let i = 1; i < dadosExistentes.length; i++) {

    if (
      normalizarCPF(
        dadosExistentes[i][1]
      ) === cpf
    ) {

      throw new Error(
        "Este CPF já está cadastrado."
      );

    }

  }

  aba.appendRow([
    gerarID(),
    cpf,
    dados.nome,
    dados.nascimento || "",
    dados.cargo || "",
    dados.setor || "",
    lojaID,
    dados.telefone || "",
    dados.email || "",
    dados.foto || "",
    "SIM",
    agora()
  ]);

  // Cria acesso do colaborador
  const usuarios =
    obterAba(ABAS.USUARIOS);

  usuarios.appendRow([
    gerarID(),
    cpf,
    gerarHash(cpf),
    dados.nome,
    "COLABORADOR",
    lojaID,
    "SIM",
    "SIM",
    agora()
  ]);

  return {
    sucesso: true,
    mensagem:
      "Colaborador cadastrado. O primeiro acesso será feito usando o CPF como senha."
  };
}


// ============================================================
// ATIVAR / DESATIVAR COLABORADOR
// ============================================================

function alterarStatusColaborador(
  token,
  id,
  ativo
) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Somente o RH pode alterar colaboradores."
    );

  }

  const aba =
    obterAba(ABAS.COLABORADORES);

  const dados =
    aba.getDataRange().getValues();

  let cpf = "";

  for (let i = 1; i < dados.length; i++) {

    if (String(dados[i][0]) === String(id)) {

      if (
        !usuarioPodeVerLoja(
          usuario,
          dados[i][6]
        )
      ) {

        throw new Error(
          "Você não possui acesso a este colaborador."
        );

      }

      aba.getRange(i + 1, 11)
        .setValue(ativo ? "SIM" : "NAO");

      cpf = normalizarCPF(dados[i][1]);

      break;
    }

  }

  if (!cpf) {
    throw new Error("Colaborador não encontrado.");
  }

  const usuarios =
    obterAba(ABAS.USUARIOS);

  const usuariosDados =
    usuarios.getDataRange().getValues();

  for (
    let i = 1;
    i < usuariosDados.length;
    i++
  ) {

    if (
      normalizarCPF(
        usuariosDados[i][1]
      ) === cpf
    ) {

      usuarios
        .getRange(i + 1, 7)
        .setValue(ativo ? "SIM" : "NAO");

      break;

    }

  }

  return {
    sucesso: true
  };
}


// ============================================================
// SOLICITAÇÃO DE DOCUMENTOS
// ============================================================

function listarDocumentos(token) {

  const usuario = obterSessao(token);

  const aba =
    obterAba(ABAS.DOCUMENTOS);

  const dados =
    aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    const linha = dados[i];

    if (
      usuario.perfil === "COLABORADOR" &&
      normalizarCPF(linha[1]) !== usuario.cpf
    ) {

      continue;

    }

    if (
      usuario.perfil !== "ADMIN" &&
      usuario.perfil !== "COLABORADOR" &&
      String(linha[3]) !== String(usuario.lojaID)
    ) {

      continue;

    }

    resultado.push({

      id: linha[0],
      cpf: linha[1],
      nome: linha[2],
      lojaID: linha[3],
      tipo: linha[4],
      descricao: linha[5],
      dataSolicitacao:
        formatarDataHora(linha[6]),
      status: linha[7],
      resposta: linha[8],
      link: linha[9],
      dataResposta:
        formatarDataHora(linha[10]),
      respondidoPor: linha[11]

    });

  }

  resultado.reverse();

  return resultado;
}


function solicitarDocumento(
  token,
  tipo,
  descricao
) {

  const usuario = obterSessao(token);

  if (usuario.perfil !== "COLABORADOR") {

    throw new Error(
      "Somente colaboradores podem solicitar documentos."
    );

  }

  const aba =
    obterAba(ABAS.DOCUMENTOS);

  aba.appendRow([
    gerarID(),
    usuario.cpf,
    usuario.nome,
    usuario.lojaID,
    tipo,
    descricao || "",
    agora(),
    "PENDENTE",
    "",
    "",
    "",
    ""
  ]);

  return {
    sucesso: true,
    mensagem:
      "Solicitação enviada ao RH."
  };
}


// ============================================================
// RESPONDER DOCUMENTO
// ============================================================

function responderDocumento(
  token,
  id,
  status,
  resposta,
  link
) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Somente o RH pode responder documentos."
    );

  }

  const aba =
    obterAba(ABAS.DOCUMENTOS);

  const dados =
    aba.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][0]) === String(id)
    ) {

      if (
        !usuarioPodeVerLoja(
          usuario,
          dados[i][3]
        )
      ) {

        throw new Error(
          "Esta solicitação pertence a outra loja."
        );

      }

      aba.getRange(i + 1, 8)
        .setValue(status);

      aba.getRange(i + 1, 9)
        .setValue(resposta || "");

      aba.getRange(i + 1, 10)
        .setValue(link || "");

      aba.getRange(i + 1, 11)
        .setValue(agora());

      aba.getRange(i + 1, 12)
        .setValue(usuario.nome);

      return {
        sucesso: true,
        mensagem:
          "Solicitação atualizada."
      };

    }

  }

  throw new Error(
    "Solicitação não encontrada."
  );
}


function contarDocumentosPendentes(token) {

  const usuario = obterSessao(token);

  const aba =
    obterAba(ABAS.DOCUMENTOS);

  const dados =
    aba.getDataRange().getValues();

  let total = 0;

  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][7]).toUpperCase() !==
      "PENDENTE"
    ) {
      continue;
    }

    if (
      !usuarioPodeVerLoja(
        usuario,
        dados[i][3]
      )
    ) {
      continue;
    }

    total++;

  }

  return total;
}


// ============================================================
// FEEDBACK
// ============================================================

function enviarFeedback(
  token,
  tipo,
  mensagem
) {

  const usuario = obterSessao(token);

  if (usuario.perfil !== "COLABORADOR") {

    throw new Error(
      "Somente colaboradores podem enviar feedback."
    );

  }

  if (!mensagem) {

    throw new Error(
      "Digite sua mensagem."
    );

  }

  const aba =
    obterAba(ABAS.FEEDBACKS);

  aba.appendRow([
    gerarID(),
    usuario.cpf,
    usuario.nome,
    usuario.lojaID,
    tipo || "Feedback",
    mensagem,
    agora(),
    "NAO"
  ]);

  return {
    sucesso: true,
    mensagem:
      "Feedback enviado ao RH."
  };
}


function listarFeedbacks(token) {

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
    obterAba(ABAS.FEEDBACKS);

  const dados =
    aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    if (
      !usuarioPodeVerLoja(
        usuario,
        dados[i][3]
      )
    ) {
      continue;
    }

    resultado.push({

      id: dados[i][0],
      nome: dados[i][2],
      tipo: dados[i][4],
      mensagem: dados[i][5],
      data: formatarDataHora(dados[i][6]),
      lido:
        String(dados[i][7]).toUpperCase() === "SIM"

    });

  }

  resultado.reverse();

  return resultado;
}


function contarFeedbacks(token) {

  const usuario = obterSessao(token);

  const aba =
    obterAba(ABAS.FEEDBACKS);

  const dados =
    aba.getDataRange().getValues();

  let total = 0;

  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][7]).toUpperCase() ===
      "SIM"
    ) {
      continue;
    }

    if (
      usuarioPodeVerLoja(
        usuario,
        dados[i][3]
      )
    ) {

      total++;

    }

  }

  return total;
}


// ============================================================
// RECONHECIMENTO
// ============================================================

function listarReconhecimentos(token) {

  const usuario = obterSessao(token);

  const aba =
    obterAba(ABAS.RECONHECIMENTOS);

  const dados =
    aba.getDataRange().getValues();

  const resultado = [];

  for (let i = 1; i < dados.length; i++) {

    if (
      !usuarioPodeVerLoja(
        usuario,
        dados[i][5]
      )
    ) {
      continue;
    }

    resultado.push({

      id: dados[i][0],
      deNome: dados[i][2],
      paraNome: dados[i][4],
      mensagem: dados[i][6],
      data: formatarDataHora(dados[i][7])

    });

  }

  resultado.reverse();

  return resultado;
}


function enviarReconhecimento(
  token,
  paraCPF,
  mensagem
) {

  const usuario = obterSessao(token);

  if (!mensagem) {
    throw new Error("Digite uma mensagem.");
  }

  const colaboradores =
    obterAba(ABAS.COLABORADORES)
      .getDataRange()
      .getValues();

  let encontrado = null;

  for (
    let i = 1;
    i < colaboradores.length;
    i++
  ) {

    if (
      normalizarCPF(
        colaboradores[i][1]
      ) === normalizarCPF(paraCPF)
    ) {

      if (
        String(colaboradores[i][6]) !==
        String(usuario.lojaID)
      ) {

        throw new Error(
          "Você só pode reconhecer colaboradores da sua loja."
        );

      }

      encontrado = colaboradores[i];

      break;

    }

  }

  if (!encontrado) {

    throw new Error(
      "Colaborador não encontrado."
    );

  }

  const aba =
    obterAba(ABAS.RECONHECIMENTOS);

  aba.appendRow([
    gerarID(),
    usuario.cpf,
    usuario.nome,
    normalizarCPF(paraCPF),
    encontrado[2],
    usuario.lojaID,
    mensagem,
    agora(),
    "SIM"
  ]);

  return {
    sucesso: true,
    mensagem:
      "Reconhecimento enviado."
  };
}


// ============================================================
// PESQUISA DE CLIMA
// ============================================================

function listarPesquisaAtiva(token) {

  const usuario = obterSessao(token);

  const aba =
    obterAba(ABAS.PESQUISAS);

  const dados =
    aba.getDataRange()
      .getValues();

  for (let i = dados.length - 1; i >= 1; i--) {

    if (
      String(dados[i][6]).toUpperCase() !==
      "SIM"
    ) {
      continue;
    }

    if (
      !usuarioPodeVerLoja(
        usuario,
        dados[i][8]
      )
    ) {
      continue;
    }

    return {

      id: dados[i][0],
      pergunta: dados[i][1],
      opcoes: [
        dados[i][2],
        dados[i][3],
        dados[i][4],
        dados[i][5]
      ].filter(String)

    };

  }

  return null;
}


function responderPesquisa(
  token,
  pesquisaID,
  resposta
) {

  const usuario = obterSessao(token);

  if (!resposta) {
    throw new Error("Selecione uma resposta.");
  }

  const respostas =
    obterAba(
      ABAS.RESPOSTAS_PESQUISA
    );

  const dados =
    respostas.getDataRange()
      .getValues();

  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][1]) ===
        String(pesquisaID) &&
      normalizarCPF(dados[i][2]) ===
        usuario.cpf
    ) {

      throw new Error(
        "Você já respondeu esta pesquisa."
      );

    }

  }

  respostas.appendRow([
    gerarID(),
    pesquisaID,
    usuario.cpf,
    usuario.nome,
    usuario.lojaID,
    resposta,
    agora()
  ]);

  return {
    sucesso: true,
    mensagem:
      "Resposta registrada."
  };
}


function criarPesquisa(
  token,
  dados
) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Somente o RH pode criar pesquisas."
    );

  }

  const lojaID =
    usuario.perfil === "ADMIN"
      ? dados.lojaID
      : usuario.lojaID;

  const aba =
    obterAba(ABAS.PESQUISAS);

  // Desativa pesquisas anteriores da mesma loja
  const valores =
    aba.getDataRange().getValues();

  for (let i = 1; i < valores.length; i++) {

    if (
      String(valores[i][8]) ===
      String(lojaID)
    ) {

      aba.getRange(i + 1, 7)
        .setValue("NAO");

    }

  }

  aba.appendRow([
    gerarID(),
    dados.pergunta,
    dados.opcao1 || "",
    dados.opcao2 || "",
    dados.opcao3 || "",
    dados.opcao4 || "",
    "SIM",
    agora(),
    lojaID,
    usuario.nome
  ]);

  return {
    sucesso: true,
    mensagem:
      "Pesquisa publicada."
  };
}


// ============================================================
// LISTAR LOJAS
// ============================================================

function listarLojas(token) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Somente o administrador pode listar todas as lojas."
    );

  }

  const aba =
    obterAba(ABAS.LOJAS);

  const dados =
    aba.getDataRange()
      .getValues();

  return dados
    .slice(1)
    .map(function(linha) {

      return {

        id: linha[0],
        nome: linha[1],
        codigo: linha[2],
        cidade: linha[3],
        ativa:
          String(linha[4]).toUpperCase() === "SIM"

      };

    });

}


// ============================================================
// LISTAR COLABORADORES PARA RECONHECIMENTO
// ============================================================

function listarColaboradoresAtivos(token) {

  const usuario = obterSessao(token);

  const colaboradores =
    listarColaboradores(token);

  return colaboradores.filter(
    function(c) {

      return (
        c.ativo &&
        normalizarCPF(c.cpf) !==
        usuario.cpf
      );

    }
  );

}


// ============================================================
// MARCAR FEEDBACK COMO LIDO
// ============================================================

function marcarFeedbackLido(
  token,
  id
) {

  const usuario = obterSessao(token);

  if (
    usuario.perfil !== "RH" &&
    usuario.perfil !== "ADMIN"
  ) {

    throw new Error(
      "Acesso negado."
    );

  }

  const aba =
    obterAba(ABAS.FEEDBACKS);

  const dados =
    aba.getDataRange()
      .getValues();

  for (let i = 1; i < dados.length; i++) {

    if (
      String(dados[i][0]) ===
      String(id)
    ) {

      if (
        !usuarioPodeVerLoja(
          usuario,
          dados[i][3]
        )
      ) {

        throw new Error(
          "Acesso negado."
        );

      }

      aba.getRange(i + 1, 8)
        .setValue("SIM");

      return true;

    }

  }

  return false;
}


// ============================================================
// FUNÇÃO DE TESTE
// ============================================================

function testarSistema() {

  const ss = obterPlanilha();

  return {

    nomePlanilha: ss.getName(),

    abas: ss
      .getSheets()
      .map(function(aba) {
        return aba.getName();
      })

  };

}

/**
 * LEADS AUTOMATION v5.5 — STABLE & CONSOLIDATED
 * Correções: CFG restaurado, funções de zebra reimplementadas, duplicatas removidas.
 */

const CFG = {
  SHEET_NAME: "pipeline",
  // ID DO MODELO (Confirme se este ID está correto para o seu template)
  TEMPLATE_ID: "1x1azS_MNjmps7ppec57NpT22S50l-KL2v5tYLZ9ed4A",

  // Colunas
  COL_CREATED_AT: "criado_em",
  COL_PROPOSAL_CODE: "codigo_prop",
  COL_PROPOSAL_VERSION: "versao_prop",
  COL_STATUS_USER: "status",
  COL_STATUS_INTERNAL: "status_interno",
  COL_EMPRESA: "empresa",
  COL_ORIGEM: "origem",
  COL_CIDADE_LOCAL: "cidade_local",
  COL_DATA_MONTAGEM: "data_montagem",
  COL_DATA_TESTES: "data_testes_ensaio",
  COL_DATA_EVENTO: "data_evento",
  COL_CLIENTE: "cliente",
  COL_CLIENTE_CONTATO: "cliente_contato_nome",
  COL_CLIENTE_EMAIL: "cliente_email",
  COL_CLIENTE_TEL: "cliente_telefone",
  COL_EVENTO_NOME: "evento_nome",
  COL_VALOR_ESTIMADO: "valor_estimado",
  COL_FORMAS_PAGAMENTO: "formas_pagamento",
  COL_DATA_EMISSAO: "proposta_data_emissao",
  COL_VALIDADE: "validade",
  COL_RESP_CONTRATANTE: "resp_contratante",
  COL_CONTATO_COMERCIAL: "contato_comercial",
  COL_COMERCIAL_EMAIL: "comercial_email",
  COL_COMERCIAL_TEL: "comercial_telefone",
  COL_FOLDER_LINK: "folder_link",
  COL_DOC_LINK: "doc_proposta_link",
  COL_PDF_LINK: "pdf_link",
  COL_EVENT_SLUG: "event_slug",
  COL_PROPOSAL_KEY: "proposal_key",
  COL_TRELLO: "trello_card_link",
  COL_ANO: "ano",
  COL_PROPOSTA_FOLDER_ID: "proposta_folder_id",
  COL_APROVACAO_FOLDER_ID: "aprovacao_folder_id",
  COL_DOC_ID: "doc_proposta_id",
  COL_OK: "ok",
  COL_OBS: "observacoes",
  COL_STATUS_LAST_PROCESSED: "status_last_processed",
  COL_LOCK_OWNER: "lock_owner",
  COL_LOCK_UNTIL: "lock_until",
  COL_LAST_RUN_ID: "last_run_id",
  COL_STATUS_UPDATED_AT: "status_updated_at",
  COL_PROPOSTA_APROVADA: "proposta_aprovada",

  SYSTEM_ADMINS: ["adm@playaudiovisual.com.br"],

  PROPOSTA_APROVADA_EDITORS: [
    "adm@playaudiovisual.com.br",
    "giuliano@playaudiovisual.com.br",
    "trodrigues@playaudiovisual.com.br",
  ],

  // Listas de Opções
  STATUS_USER_OPTIONS: [
    "Novo",
    "Em negociação",
    "Nova versão",
    "Aprovado",
    "Gerar PDF",
    "Não aprovado",
    "Cancelado",
  ],
  INTERNAL_STATUS_OPTIONS: [
    "Proposta - Gerar",
    "proposta_em_edicao",
    "transformar_em_pdf",
  ],
  PAGAMENTO_OPTIONS: [
    "50% na contratação + 50% para 10 dias apos evento",
    "50% na contratação + 50% para 30 dias apos o evento",
    "100% na contratação",
  ],
  EMPRESAS: ["Play Audiovisual", "Play Projeções", "Play Total"],
  ORIGENS: [
    "Site",
    "Cliente",
    "Telefone",
    "WhatsApp",
    "Integração",
    "Teatro Santander",
    "Teatro Iguatemi",
    "Teatro Multiplan",
    "Teatro Villa Lobos",
    "Rooftop 033",
    "Teatro + Rooftop 033",
  ],
  CONTATOS_COMERCIAIS: ["Giuliano", "Thayany", "Bruna", "Janaína", "Márcio"],

  // Dados Comerciais
  CONTATO2INFO: {
    Giuliano: {
      email: "giuliano@playaudiovisual.com.br",
      telefone: "11 94791 4699",
    },
    Thayany: {
      email: "trodrigues@playaudiovisual.com.br",
      telefone: "11 94001 9066",
    },
    Bruna: {
      email: "bruna@playaudiovisual.com.br",
      telefone: "11 94031 1598",
    },
    Janaína: {
      email: "janaina@playaudiovisual.com.br",
      telefone: "11 94001-9355",
    },
    Márcio: {
      email: "marciocosta@playaudiovisual.com.br",
      telefone: "11 97212 0345",
    },
  },

  // Configurações de Sistema (RESTAURADAS)
  SYSTEM_HEADERS: [
    "criado_em",
    "codigo_prop",
    "versao_prop",
    "status_interno",
    "proposta_data_emissao",
    "comercial_email",
    "comercial_telefone",
    "proposta_folder_id",
    "aprovacao_folder_id",
    "doc_proposta_id",
    "proposal_key",
    "event_slug",
    "ano",
    "ok",
    "status_last_processed",
    "lock_owner",
    "lock_until",
    "last_run_id",
    "status_updated_at",
  ],

  HIDDEN_HEADERS: [
    "status_interno",
    "comercial_email",
    "comercial_telefone",
    "proposta_folder_id",
    "aprovacao_folder_id",
    "doc_proposta_id",
    "proposal_key",
    "status_last_processed",
    "lock_owner",
    "lock_until",
    "last_run_id",
    "status_updated_at",
    "snapshot_json",
  ],

  REQUIRED_GATING_HEADERS: [
    "origem",
    "cliente",
    "cliente_contato_nome",
    "empresa",
  ],
  REQUIRED_USER_HEADERS: [
    "cidade_local",
    "data_montagem",
    "data_testes_ensaio",
    "data_evento",
    "cliente_email",
    "cliente_telefone",
    "evento_nome",
    "valor_estimado",
    "formas_pagamento",
  ],

  PROP_KEY: "NEXT_PROPOSAL_CODE",
  PROP_START_AT: 10000,
};

// Paleta ÚNICA de cores (somente para a coluna STATUS)
const STATUS_BG_MAP = {
  Novo: "#c5e7ff", // azul light
  "Em negociação": "#fef9c0", // amarelo light
  "Nova versão": "#f1d7f5", // roxo light
  Aprovado: "#c8f6cb", // verde light
  "Gerar PDF": "#decff4", // lilás light
  "Não aprovado": "#fbce85", // laranja light
  Cancelado: "#fe97a7", // vermelho light
};

const N8N = {
  WEBHOOK_NOVA_PROPOSTA:
    "https://automacao.playaudiovisual.com.br/webhook/proposta/novo",

  WEBHOOK_ONEDIT: "https://automacao.playaudiovisual.com.br/webhook/onEdit",

  WEBHOOK_GERAR_PDF:
    "https://automacao.playaudiovisual.com.br/webhook/proposta/gerar-pdf",

  WEBHOOK_APROVADO:
    "https://automacao.playaudiovisual.com.br/webhook/proposta/aprovado",

  WEBHOOK_NOVA_VERSAO:
    "https://automacao.playaudiovisual.com.br/webhook/onEdit",

  WEBHOOK_CANCELADO:
    "https://automacao.playaudiovisual.com.br/webhook/proposta/cancelado",
};
const LOCK_TTL_MIN = 5;
const APPROVAL_WEBHOOK_PROP_PREFIX = "approval_webhook_sent";

// ============== GATILHO SIMPLES (UI) ==============
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_NAME) return;
    const row = e.range.getRow();
    if (row === 1) return; // pula cabeçalho

    const h = getHeaderMap_();
    const editedCol = Object.keys(h).find(
      (key) => h[key] === e.range.getColumn(),
    );
    if (!editedCol) return;

    // Bloqueia edição em colunas de sistema
    if (CFG.SYSTEM_HEADERS.includes(editedCol)) {
      e.range.setValue(e.oldValue || "");
      toast_("Campo sistema. Edição bloqueada.");
      return;
    }

    // Bloqueia edição na coluna proposta_aprovada para não autorizados
    if (editedCol === CFG.COL_PROPOSTA_APROVADA) {
      if (!enforcePropostaAprovadaPermission_(e)) return;
    }

    // Quando editar cliente_telefone, normaliza
    if (editedCol === CFG.COL_CLIENTE_TEL) {
      normalizePhoneCell_(row, CFG.COL_CLIENTE_TEL);
    }

    // Bloqueia mudança manual de status restrito (Aprovado/Cancelado)
    if (editedCol === CFG.COL_STATUS_USER) {
      if (!enforceStatusPermissions_(e, editedCol)) return;
    }

    paintRowsFromRange_(e.range); // pinta todas as linhas afetadas (cola/preenchimento em bloco)
  } catch (err) {
    Logger.log(err);
  }
}

// ============== GATILHO INSTALÁVEL (LÓGICA) ==============
function onEditInstallable(e) {
  console.log(">>> INSTALLABLE TRIGGER <<<");
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CFG.SHEET_NAME) return;
    const row = e.range.getRow();
    if (row === 1) return;

    const h = getHeaderMap_();
    const editedCol = Object.keys(h).find(
      (key) => h[key] === e.range.getColumn(),
    );

    // Se não achou a coluna editada, sai
    if (!editedCol) return;

    // Bloqueia mudança manual de status restrito ANTES de qualquer lógica e antes de webhook
    if (editedCol === CFG.COL_STATUS_USER) {
      if (!enforceStatusPermissions_(e, editedCol)) return;
    }

    // Bloqueia edição na coluna proposta_aprovada para não autorizados
    if (editedCol === CFG.COL_PROPOSTA_APROVADA) {
      if (!enforcePropostaAprovadaPermission_(e)) return;
    }

    // 1. AUTO-FILL (gera criado_em, codigo_prop, versao_prop, status, validade)
    const generatedId = maybeAutofillRow_(row);

    // 2. Garante que TODAS as colunas obrigatórias estejam preenchidas
    const requiredKeys = getRequiredLeadKeys_();

    const stableRowValues = ensureRowStable_(row, {
      requiredCols: requiredKeys,
      attempts: 6, // lê até 6 vezes para evitar condição de corrida
      delayMs: 250, // 0.25s entre leituras
    });

    const payload = buildN8NPayload_(row, stableRowValues);

    // Se o código foi recém-gerado, garante que está no payload
    if ((!payload.codigo_prop || payload.codigo_prop === "") && generatedId) {
      payload.codigo_prop = generatedId;
    }

    const missing = requiredKeys.filter((k) => isBlankValue_(payload[k]));
    const isDataComplete = missing.length === 0;

    // 3. Quando o usuário escolhe o contato_comercial, preenche e formata o comercial_telefone
    if (editedCol === CFG.COL_CONTATO_COMERCIAL) {
      fillComercialFromContato_(row);
      SpreadsheetApp.flush();
    }

    // 4. Lógica de disparo para o N8N
    const statusNow = String(
      sh.getRange(row, h[CFG.COL_STATUS_USER]).getValue() || "",
    ).trim();
    const statusKey = normalizeStatusKey_(statusNow);
    const oldStatusKey =
      editedCol === CFG.COL_STATUS_USER
        ? normalizeStatusKey_(
            typeof e.oldValue !== "undefined" ? e.oldValue : "",
          )
        : "";
    const isManualStatusTransition =
      editedCol === CFG.COL_STATUS_USER && oldStatusKey !== statusKey;

    const shouldMarkStatusUpdate =
      editedCol === CFG.COL_STATUS_USER ||
      (statusKey === "novo" &&
        (requiredKeys.includes(editedCol) ||
          editedCol === CFG.COL_STATUS_USER));

    if (shouldMarkStatusUpdate) {
      markStatusUpdated_(row);
    }

    // Só queremos disparar automático no status "Novo"
    if (statusKey === "novo") {
      if (!isDataComplete) {
        console.log(
          `[SKIP] Linha ${row} incompleta. Faltando: ${missing.join(", ")}`,
        );
        return;
      }
      if (!payload.codigo_prop) return;

      const isRelevantEdit =
        requiredKeys.includes(editedCol) || editedCol === CFG.COL_STATUS_USER;

      if (!isRelevantEdit) return;
    }

    const statusToWebhook = {
      novo: N8N.WEBHOOK_NOVA_PROPOSTA,
      "nova versao": N8N.WEBHOOK_NOVA_VERSAO,
      "gerar pdf": N8N.WEBHOOK_GERAR_PDF,
      aprovado: N8N.WEBHOOK_APROVADO,
      cancelado: N8N.WEBHOOK_CANCELADO,
    };

    // SEMPRE repinta as linhas afetadas (visual)
    paintRowsFromRange_(e.range);

    const webhookUrl = statusToWebhook[statusKey];

    if (webhookUrl && (isDataComplete || statusKey !== "novo")) {
      const updAt = new Date(
        sh.getRange(row, h[CFG.COL_STATUS_UPDATED_AT]).getValue() || 0,
      ).getTime();
      const sendDecision = shouldSendWebhook_(row, statusKey, {
        updatedAtMs: updAt,
        editedCol,
        isManualStatusTransition,
        payload,
      });

      if (sendDecision.canSend) {
        const runId = Utilities.getUuid();
        payload.__run_uuid = runId;

        console.log(`🚀 ENVIANDO (${statusNow}): ID=${payload.codigo_prop}`);
        const sendResult = postToN8N_(payload, webhookUrl);

        if (sendResult.ok) {
          registerWebhookSuccess_(row, statusKey, runId, payload);
        } else {
          releaseRowLock_(row);
          console.error(
            `❌ FALHA (${statusNow}): HTTP=${sendResult.statusCode || 0} ${sendResult.error || sendResult.body || ""}`,
          );
        }
      } else {
        console.log(
          `⛔ BLOQUEADO (${statusNow}): linha ${row}. Motivo: ${sendDecision.reason}`,
        );
      }
    }

    if (editedCol === CFG.COL_DATA_EMISSAO) setAnoFromEmissao_(row);
  } catch (err) {
    console.error(err);
  }
}

// ============== LÓGICA DE NEGÓCIO ==============

function maybeAutofillRow_(row) {
  const sh = getSheet_();
  let generatedId = null;

  const already = [CFG.COL_PROPOSAL_CODE, CFG.COL_CREATED_AT].some((name) => {
    try {
      return !!sh.getRange(row, col_(name)).getValue();
    } catch (_) {
      return false;
    }
  });
  if (already) {
    try {
      generatedId = sh.getRange(row, col_(CFG.COL_PROPOSAL_CODE)).getValue();
    } catch (_) {}
    return generatedId;
  }

  const gatingOk = CFG.REQUIRED_GATING_HEADERS.every((name) => {
    try {
      const v = sh.getRange(row, col_(name)).getValue();
      return v !== "" && v !== null;
    } catch (_) {
      return false;
    }
  });
  if (!gatingOk) return null;

  const t = now_();
  sh.getRange(row, col_(CFG.COL_CREATED_AT)).setValue(t);

  withLock_(() => {
    const cell = sh.getRange(row, col_(CFG.COL_PROPOSAL_CODE));
    if (String(cell.getValue() || "").trim()) {
      generatedId = cell.getValue();
      return;
    }
    const props = PropertiesService.getDocumentProperties();
    let n = parseInt(props.getProperty(CFG.PROP_KEY) || CFG.PROP_START_AT, 10);
    if (isNaN(n) || n < CFG.PROP_START_AT) n = CFG.PROP_START_AT;
    cell.setValue(n);
    generatedId = n;
    props.setProperty(CFG.PROP_KEY, String(n + 1));
  });

  sh.getRange(row, col_(CFG.COL_PROPOSAL_VERSION)).setValue("A");
  sh.getRange(row, col_(CFG.COL_STATUS_USER)).setValue("Novo");

  paintRowColors_(row); // Pinta imediatamente

  sh.getRange(row, col_(CFG.COL_DATA_EMISSAO)).setValue(t);
  setAnoFromEmissao_(row);

  SpreadsheetApp.flush();
  return generatedId;
}

function getRequiredLeadKeys_() {
  return [
    CFG.COL_EMPRESA,
    CFG.COL_CONTATO_COMERCIAL,
    CFG.COL_ORIGEM,
    CFG.COL_CIDADE_LOCAL,
    CFG.COL_DATA_MONTAGEM,
    CFG.COL_DATA_TESTES,
    CFG.COL_DATA_EVENTO,
    CFG.COL_CLIENTE,
    CFG.COL_CLIENTE_CONTATO,
    CFG.COL_CLIENTE_EMAIL,
    CFG.COL_CLIENTE_TEL,
    CFG.COL_EVENTO_NOME,
    CFG.COL_VALOR_ESTIMADO,
    CFG.COL_FORMAS_PAGAMENTO,
  ];
}

function fillComercialFromContato_(row) {
  const sh = getSheet_();
  const cContato = col_(CFG.COL_CONTATO_COMERCIAL);
  const val = (sh.getRange(row, cContato).getValue() || "").toString().trim();
  const info = CFG.CONTATO2INFO[val];
  if (!info) return;
  try {
    sh.getRange(row, col_(CFG.COL_COMERCIAL_EMAIL)).setValue(info.email);
  } catch (_) {}
  try {
    sh.getRange(row, col_(CFG.COL_COMERCIAL_TEL)).setValue(
      formatPhone_(info.telefone),
    );
  } catch (_) {}
}

function arraysEqual(a, b) {
  if (!a || !b) return false;
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (String(a[i] || "") !== String(b[i] || "")) return false;
  }
  return true;
}

function ensureRowStable_(row, options = {}) {
  const sh = getSheet_();
  const attempts = options.attempts || 6;
  const delay = options.delayMs || 200;
  const h = getHeaderMap_();
  const lastCol = sh.getLastColumn();
  let prev = null;
  for (let i = 0; i < attempts; i++) {
    const vals = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    // if we have required columns configured, check all are present
    if (options.requiredCols && options.requiredCols.length > 0) {
      const ok = options.requiredCols.every((colName) => {
        const c = h[colName];
        if (!c) return false;
        const v = vals[c - 1];
        return v !== "" && v !== null && v !== undefined;
      });
      if (ok) return vals;
    }
    // check for stability (same result in two consecutive reads)
    if (prev && arraysEqual(prev, vals)) return vals;
    prev = vals;
    // give the sheet a moment to settle and any other triggers to finish
    SpreadsheetApp.flush();
    Utilities.sleep(delay);
  }
  // return last read even if not stable
  return prev;
}

// Alter buildN8NPayload_ to accept optional rowValues
function buildN8NPayload_(row, rowValues) {
  const sh = getSheet_();
  const h = getHeaderMap_();
  const values =
    rowValues || sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const obj = {};
  Object.keys(h).forEach((k) => (obj[k] = values[h[k] - 1]));

  const contato = (obj[CFG.COL_CONTATO_COMERCIAL] || "").toString().trim();
  const info = CFG.CONTATO2INFO[contato];
  if (info) {
    obj[CFG.COL_COMERCIAL_EMAIL] = info.email;
    obj[CFG.COL_COMERCIAL_TEL] = info.telefone;
  }

  obj.__rowNumber = row;
  obj.__sheetName = sh.getName();
  obj.__edited_by = Session.getActiveUser().getEmail();
  obj.__run_uuid = Utilities.getUuid();
  return obj;
}

// Add aliases for backward compatibility (in case other scripts expect these names)
function backfillAno() {
  if (typeof backfillAno_ === "function") return backfillAno_();
  Logger.log("backfillAno_ not found");
}
function backfillComercialAll() {
  if (typeof backfillComercialAll_ === "function")
    return backfillComercialAll_();
  Logger.log("backfillComercialAll_ not found");
}

// ============== HELPERS ==============

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_NAME);
}
function getHeaderMap_() {
  const sh = getSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h && typeof h === "string") map[h.trim()] = i + 1;
  });
  return map;
}
function col_(name) {
  const map = getHeaderMap_();
  const c = map[name];
  if (!c) throw new Error("Coluna não encontrada: " + name);
  return c;
}
function now_() {
  return new Date();
}
function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(45000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
function formatPhone_(raw) {
  if (!raw) return "";
  const d = String(raw).replace(/\D+/g, "").slice(0, 11);
  if (d.length === 11) return `${d.slice(0, 2)} ${d.slice(2, 7)} ${d.slice(7)}`;
  if (d.length === 10) return `${d.slice(0, 2)} ${d.slice(2, 6)} ${d.slice(6)}`;
  return d;
}

function normalizeStatusKey_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function parseDateValueMs_(value) {
  if (!value) return 0;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    const ms = value.getTime();
    return isNaN(ms) ? 0 : ms;
  }
  const ms = new Date(value).getTime();
  return isNaN(ms) ? 0 : ms;
}

function isBlankValue_(value) {
  if (typeof value === "string") return value.trim() === "";
  return value === "" || value === null || value === undefined;
}

function getApprovalWebhookKey_(row, payload) {
  const code = String(payload?.[CFG.COL_PROPOSAL_CODE] || "").trim();
  const version = String(payload?.[CFG.COL_PROPOSAL_VERSION] || "A")
    .trim()
    .toUpperCase();
  const proposalKey = String(payload?.[CFG.COL_PROPOSAL_KEY] || "").trim();

  if (code) return `${APPROVAL_WEBHOOK_PROP_PREFIX}:${code}:${version}`;
  if (proposalKey)
    return `${APPROVAL_WEBHOOK_PROP_PREFIX}:${proposalKey}:${version}`;
  return `${APPROVAL_WEBHOOK_PROP_PREFIX}:row:${row}:${version}`;
}

function hasApprovalWebhookBeenSent_(row, payload) {
  const key = getApprovalWebhookKey_(row, payload);
  return !!PropertiesService.getDocumentProperties().getProperty(key);
}

function markApprovalWebhookSent_(row, payload, meta = {}) {
  const key = getApprovalWebhookKey_(row, payload);
  const record = {
    row,
    codigo_prop: String(payload?.[CFG.COL_PROPOSAL_CODE] || ""),
    versao_prop: String(payload?.[CFG.COL_PROPOSAL_VERSION] || ""),
    proposal_key: String(payload?.[CFG.COL_PROPOSAL_KEY] || ""),
    sent_at: meta.sentAt || new Date().toISOString(),
    by: meta.by || "",
    run_id: meta.runId || "",
  };

  PropertiesService.getDocumentProperties().setProperty(
    key,
    JSON.stringify(record),
  );
  return key;
}

function clearApprovalWebhookSent_(row, payload) {
  const key = getApprovalWebhookKey_(row, payload);
  PropertiesService.getDocumentProperties().deleteProperty(key);
  return key;
}

function normalizePhoneCell_(row, colName) {
  const sh = getSheet_();
  const c = col_(colName);
  const raw = sh.getRange(row, c).getValue();
  sh.getRange(row, c).setValue(formatPhone_(raw));
}

function postToN8N_(payload, url) {
  if (!url) {
    return { ok: false, statusCode: 0, error: "missing_url", body: "" };
  }

  try {
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const body = response.getContentText() || "";
    const ok = statusCode >= 200 && statusCode < 300;

    if (ok) {
      console.log(`Enviado para N8N: ${url} (HTTP ${statusCode})`);
    } else {
      console.error(`Falha N8N: ${url} (HTTP ${statusCode}) ${body}`);
    }

    return { ok, statusCode, body, error: "" };
  } catch (e) {
    console.error(`Erro envio N8N: ${e.message}`);
    return {
      ok: false,
      statusCode: 0,
      error: e.message || String(e),
      body: "",
    };
  }
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, "Leads", 3);
}
function isColumnHidden_(sh, colIndex) {
  try {
    return sh.isColumnHiddenByUser(colIndex);
  } catch (_) {
    return false;
  }
}
function markStatusUpdated_(row) {
  try {
    getSheet_()
      .getRange(row, col_(CFG.COL_STATUS_UPDATED_AT))
      .setValue(new Date().toISOString());
  } catch (_) {}
}
function acquireRowLock_(row, updatedAtMs) {
  const sh = getSheet_();
  const h = getHeaderMap_();
  const colOwner = h[CFG.COL_LOCK_OWNER];
  const colUntil = h[CFG.COL_LOCK_UNTIL];
  const colLastProcessed = h[CFG.COL_STATUS_LAST_PROCESSED];
  const now = Date.now();
  const until = colUntil
    ? Number(sh.getRange(row, colUntil).getValue() || 0)
    : 0;
  const lastProcessedMs = colLastProcessed
    ? parseDateValueMs_(sh.getRange(row, colLastProcessed).getValue())
    : 0;

  if (until && now < until) {
    // Lock ativo: bloqueia incondicionalmente para evitar disparos duplos
    // durante o processamento do N8N (que pode levar até 30s).
    return false;
  }

  if (colOwner) {
    sh.getRange(row, colOwner).setValue(Session.getActiveUser().getEmail());
  }
  if (colUntil) {
    sh.getRange(row, colUntil).setValue(now + LOCK_TTL_MIN * 60000);
  }
  return true;
}

function releaseRowLock_(row) {
  const sh = getSheet_();
  const h = getHeaderMap_();

  try {
    if (h[CFG.COL_LOCK_OWNER]) {
      sh.getRange(row, h[CFG.COL_LOCK_OWNER]).setValue("");
    }
  } catch (_) {}

  try {
    if (h[CFG.COL_LOCK_UNTIL]) {
      sh.getRange(row, h[CFG.COL_LOCK_UNTIL]).setValue("");
    }
  } catch (_) {}
}

function registerWebhookSuccess_(row, statusKey, runId, payload) {
  const sh = getSheet_();
  const h = getHeaderMap_();
  const processedAt = new Date().toISOString();

  try {
    if (h[CFG.COL_LAST_RUN_ID]) {
      sh.getRange(row, h[CFG.COL_LAST_RUN_ID]).setValue(runId);
    }
  } catch (_) {}

  try {
    if (h[CFG.COL_STATUS_LAST_PROCESSED]) {
      sh.getRange(row, h[CFG.COL_STATUS_LAST_PROCESSED]).setValue(processedAt);
    }
  } catch (_) {}

  if (statusKey === "aprovado") {
    markApprovalWebhookSent_(row, payload, {
      sentAt: processedAt,
      by: payload?.__edited_by || "",
      runId,
    });
  }
}

function shouldSendWebhook_(row, statusKey, options = {}) {
  const updatedAtMs = Number(options.updatedAtMs || 0);
  const editedCol = String(options.editedCol || "");
  const isManualStatusTransition = !!options.isManualStatusTransition;
  const payload = options.payload || {};
  const isManualStatusChange = editedCol === CFG.COL_STATUS_USER;

  // Mantém o comportamento do arquivo antigo para Novo/Nova versão/Gerar PDF/Cancelado:
  // mudança manual de status passa direto; demais casos usam lock de linha.
  if (statusKey !== "aprovado") {
    if (isManualStatusChange) {
      return { canSend: true, reason: "" };
    }

    if (!acquireRowLock_(row, updatedAtMs)) {
      return {
        canSend: false,
        reason: "linha ainda está em lock da execução anterior",
      };
    }

    return { canSend: true, reason: "" };
  }

  if (statusKey === "aprovado") {
    if (editedCol !== CFG.COL_STATUS_USER) {
      return {
        canSend: false,
        reason: "Aprovado só dispara em edição manual da coluna status",
      };
    }

    if (!isManualStatusTransition) {
      return {
        canSend: false,
        reason: "o status já estava Aprovado",
      };
    }

    if (hasApprovalWebhookBeenSent_(row, payload)) {
      return {
        canSend: false,
        reason: "esta versão já teve o webhook de Aprovado enviado",
      };
    }
  }

  if (!acquireRowLock_(row, updatedAtMs)) {
    return {
      canSend: false,
      reason: "linha ainda está em lock da execução anterior",
    };
  }

  return { canSend: true, reason: "" };
}

function processPendingNovoRows() {
  const sh = getSheet_();
  if (!sh) return;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sh.getLastColumn();
  const requiredKeys = getRequiredLeadKeys_();

  for (let row = 2; row <= lastRow; row++) {
    try {
      const rowValues = sh.getRange(row, 1, 1, lastCol).getValues()[0];
      const payload = buildN8NPayload_(row, rowValues);
      const statusKey = normalizeStatusKey_(payload[CFG.COL_STATUS_USER]);

      if (statusKey !== "novo") continue;
      if (!payload.codigo_prop) continue;

      const missing = requiredKeys.filter((k) => isBlankValue_(payload[k]));
      if (missing.length > 0) continue;

      let updatedAtMs = parseDateValueMs_(payload[CFG.COL_STATUS_UPDATED_AT]);
      const lastProcessedMs = parseDateValueMs_(
        payload[CFG.COL_STATUS_LAST_PROCESSED],
      );

      // Cooldown absoluto: não re-envia se o webhook já foi enviado nos últimos 10 minutos,
      // independente de edições na linha. Evita disparos duplos durante processamento do N8N.
      const COOLDOWN_MS = 10 * 60 * 1000;
      if (lastProcessedMs && Date.now() - lastProcessedMs < COOLDOWN_MS) {
        continue;
      }

      if (lastProcessedMs && (!updatedAtMs || updatedAtMs <= lastProcessedMs)) {
        continue;
      }

      if (!updatedAtMs) {
        markStatusUpdated_(row);
        payload[CFG.COL_STATUS_UPDATED_AT] = new Date().toISOString();
        updatedAtMs = parseDateValueMs_(payload[CFG.COL_STATUS_UPDATED_AT]);
      }

      const sendDecision = shouldSendWebhook_(row, statusKey, {
        updatedAtMs,
        payload,
      });

      if (!sendDecision.canSend) continue;

      const runId = Utilities.getUuid();
      payload.__run_uuid = runId;

      console.log(
        `[NOVO/PENDENTE] ENVIANDO linha ${row}: ID=${payload.codigo_prop}`,
      );
      const sendResult = postToN8N_(payload, N8N.WEBHOOK_NOVA_PROPOSTA);

      if (sendResult.ok) {
        registerWebhookSuccess_(row, statusKey, runId, payload);
      } else {
        releaseRowLock_(row);
        console.error(
          `[NOVO/PENDENTE] FALHA linha ${row}: HTTP=${sendResult.statusCode || 0} ${sendResult.error || sendResult.body || ""}`,
        );
      }
    } catch (err) {
      console.error(`[NOVO/PENDENTE] Erro linha ${row}: ${err}`);
    }
  }
}

function setAnoFromEmissao_(row) {
  try {
    const sh = getSheet_();
    if (!row || row < 2) return;

    const h = getHeaderMap_();
    const colAno = col_(CFG.COL_ANO);
    if (!colAno) return;

    const anoCell = sh.getRange(row, colAno);
    const anoValue = String(anoCell.getValue() || "").trim();
    if (anoValue !== "") return; // não sobrescrever

    // Prioridade: data_emissao > criado_em > data_evento
    const candidates = [
      col_(CFG.COL_DATA_EMISSAO),
      col_(CFG.COL_CREATED_AT),
      col_(CFG.COL_DATA_EVENTO),
    ];

    for (let c of candidates) {
      if (!c) continue;
      const raw = sh.getRange(row, c).getValue();
      if (raw === undefined || raw === null || String(raw).trim() === "")
        continue;

      let year = null;
      if (typeof toYearSmart === "function") {
        year = toYearSmart(raw);
      } else {
        // fallback simples
        const d = new Date(raw);
        if (!isNaN(d)) year = d.getFullYear();
      }
      if (year && !isNaN(year)) {
        anoCell.setValue(Number(year));
        return;
      }
    }
  } catch (_) {
    // swallow errors para não quebrar outros processos
  }
}

function backfillAno_() {
  const sh = getSheet_();
  const lastRow = Math.max(1, sh.getLastRow());
  if (lastRow < 2) return;
  for (let r = 2; r <= lastRow; r++) {
    try {
      setAnoFromEmissao_(r);
    } catch (_) {}
  }
}

function backfillComercialAll_() {
  const sh = getSheet_();
  const lastRow = Math.max(1, sh.getLastRow());
  if (lastRow < 2) return;
  for (let r = 2; r <= lastRow; r++) {
    try {
      fillComercialFromContato_(r);
    } catch (_) {}
  }
}

// ==============================
// PERMISSÕES PARA STATUS (UI)
// ==============================

function getActiveEditorEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || "")
      .toLowerCase()
      .trim();
  } catch (_) {
    return "";
  }
}

function isSystemAdmin_(email) {
  const em = String(email || "")
    .toLowerCase()
    .trim();
  if (!em) return false;

  // Owner da planilha (recomendado manter com permissão)
  try {
    const owner = SpreadsheetApp.getActiveSpreadsheet().getOwner();
    const ownerEmail = owner
      ? String(owner.getEmail() || "")
          .toLowerCase()
          .trim()
      : "";
    if (ownerEmail && em === ownerEmail) return true;
  } catch (_) {}

  // SYSTEM_ADMINS (já existe no CFG)
  const admins = (CFG.SYSTEM_ADMINS || []).map((x) =>
    String(x || "")
      .toLowerCase()
      .trim(),
  );
  return admins.includes(em);
}

function canSetRestrictedStatus_(email) {
  const em = String(email || "")
    .toLowerCase()
    .trim();
  if (!em) return false;

  // Admin/owner sempre pode (recomendado)
  if (isSystemAdmin_(em)) return true;

  // Somente Giuliano + Thayany (email vem do seu CFG.CONTATO2INFO)
  const allowed = [
    CFG.CONTATO2INFO?.Giuliano?.email,
    CFG.CONTATO2INFO?.Thayany?.email,
  ]
    .filter(Boolean)
    .map((x) => String(x).toLowerCase().trim());

  return allowed.includes(em);
}

function canEditPropostaAprovada_(email) {
  const em = String(email || "")
    .toLowerCase()
    .trim();
  if (!em) return false;
  const allowed = (CFG.PROPOSTA_APROVADA_EDITORS || []).map((x) =>
    String(x || "")
      .toLowerCase()
      .trim(),
  );
  return allowed.includes(em);
}

/**
 * Bloqueia edição na coluna "proposta_aprovada" para usuários não autorizados.
 * Reverte para o valor anterior e exibe toast.
 *
 * Retorna true se pode continuar, false se bloqueou.
 */
function enforcePropostaAprovadaPermission_(e) {
  try {
    if (!e || !e.range) return true;
    const email = getActiveEditorEmail_();
    if (canEditPropostaAprovada_(email)) return true;

    // Sem permissão: reverte
    if (typeof e.oldValue !== "undefined") {
      e.range.setValue(e.oldValue || "");
    } else {
      e.range.setValue("");
    }

    const who = email ? email : "usuário não identificado";
    toast_(`Sem permissão para editar "proposta_aprovada". (${who})`);
    return false;
  } catch (err) {
    Logger.log(err);
    return true;
  }
}

/**
 * Bloqueia mudança manual de status restrito ("Aprovado" / "Cancelado")
 * para usuários não autorizados. Reverte para o valor anterior.
 *
 * Retorna true se pode continuar, false se bloqueou.
 */
function enforceStatusPermissions_(e, editedCol) {
  try {
    if (!e || !e.range) return true;
    if (editedCol !== CFG.COL_STATUS_USER) return true;

    // OBS: aqui é manual (UI). Para atualizações por API/workflow, isso normalmente nem dispara.
    const newStatus = String(e.range.getValue() || "").trim();
    const key = newStatus.toLowerCase();

    const isRestricted = key === "aprovado" || key === "cancelado";
    if (!isRestricted) return true;

    const email = getActiveEditorEmail_();
    if (canSetRestrictedStatus_(email)) return true;

    // Sem permissão: reverte
    if (typeof e.oldValue !== "undefined") {
      e.range.setValue(e.oldValue || "");
    } else {
      // Em alguns cenários (colar intervalos), oldValue pode não existir.
      // Preferimos limpar para não deixar status restrito setado sem permissão.
      e.range.setValue("");
    }

    // Reaplica cor da linha/status
    paintRowColors_(e.range.getRow());

    const who = email ? email : "usuário não identificado";
    toast_(`Sem permissão para definir "${newStatus}". (${who})`);
    return false;
  } catch (err) {
    Logger.log(err);
    return true; // falha “safe”: não derruba o resto do fluxo
  }
}

/**
 * Reaplica as cores (zebra + status) para todas as linhas afetadas pelo range editado.
 * IMPORTANTE: isso resolve colar/preencher múltiplas linhas de uma vez.
 */
function paintRowsFromRange_(range) {
  if (!range) return;

  const sh = range.getSheet();
  if (sh.getName() !== CFG.SHEET_NAME) return;

  const startRow = range.getRow();
  const numRows = range.getNumRows ? range.getNumRows() : 1;

  // Não pinta header
  const from = Math.max(2, startRow);
  const to = startRow + numRows - 1;

  for (let r = from; r <= to; r++) {
    paintRowColors_(r);
  }
}

// ======================================================================
// INFRAESTRUTURA, SETUP E VISUAL (Restaurado)
// ======================================================================

function ensureNextCodeProperty_() {
  const props = PropertiesService.getDocumentProperties();
  if (!props.getProperty(CFG.PROP_KEY))
    props.setProperty(CFG.PROP_KEY, String(CFG.PROP_START_AT));
}

/**
 * Retorna o Range dinâmico de uma coluna da aba LISTAS (A=empresa, B=contato, C=origem).
 * Descobre automaticamente a última linha preenchida, então novos itens
 * adicionados na aba LISTAS aparecem no dropdown sem precisar rodar o setup novamente.
 */
function getListaRange_(colLetter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const listasSh = ss.getSheetByName("LISTAS");
  if (!listasSh) throw new Error("Aba 'LISTAS' não encontrada.");
  const colIndex = colLetter.toUpperCase().charCodeAt(0) - 64; // A→1, B→2, C→3
  const lastRow = listasSh.getLastRow();
  const endRow = Math.max(2, lastRow);
  return listasSh.getRange(2, colIndex, endRow - 1, 1);
}

function applyValidations_() {
  const sh = getSheet_();
  const lastRow = Math.max(2, sh.getMaxRows());
  if (lastRow < 2) return;

  const dvList = (v) =>
    SpreadsheetApp.newDataValidation()
      .requireValueInList(v, true)
      .setAllowInvalid(false)
      .build();

  const dvRange = (range) =>
    SpreadsheetApp.newDataValidation()
      .requireValueInRange(range, true)
      .setAllowInvalid(false)
      .build();

  try {
    sh.getRange(2, col_(CFG.COL_STATUS_USER), lastRow - 1).setDataValidation(
      dvList(CFG.STATUS_USER_OPTIONS),
    );
  } catch (_) {}
  try {
    sh.getRange(
      2,
      col_(CFG.COL_STATUS_INTERNAL),
      lastRow - 1,
    ).setDataValidation(dvList(CFG.INTERNAL_STATUS_OPTIONS));
  } catch (_) {}
  try {
    // EMPRESA → LISTAS!A2:A (dinâmico)
    sh.getRange(2, col_(CFG.COL_EMPRESA), lastRow - 1).setDataValidation(
      dvRange(getListaRange_("A")),
    );
  } catch (_) {}
  try {
    // ORIGEM → LISTAS!C2:C (dinâmico)
    sh.getRange(2, col_(CFG.COL_ORIGEM), lastRow - 1).setDataValidation(
      dvRange(getListaRange_("C")),
    );
  } catch (_) {}
  try {
    sh.getRange(
      2,
      col_(CFG.COL_FORMAS_PAGAMENTO),
      lastRow - 1,
    ).setDataValidation(dvList(CFG.PAGAMENTO_OPTIONS));
  } catch (_) {}
  try {
    // CONTATO COMERCIAL → LISTAS!B2:B (dinâmico)
    sh.getRange(
      2,
      col_(CFG.COL_CONTATO_COMERCIAL),
      lastRow - 1,
    ).setDataValidation(dvRange(getListaRange_("B")));
  } catch (_) {}

  const letters = Array.from({ length: 26 }, (_, i) =>
    String.fromCharCode(65 + i),
  );
  try {
    sh.getRange(
      2,
      col_(CFG.COL_PROPOSAL_VERSION),
      lastRow - 1,
    ).setDataValidation(dvList(letters));
  } catch (_) {}

  try {
    sh.getRange(2, col_(CFG.COL_CLIENTE_EMAIL), lastRow - 1).setDataValidation(
      SpreadsheetApp.newDataValidation()
        .requireTextIsEmail()
        .setAllowInvalid(true)
        .build(),
    );
  } catch (_) {}

  try {
    sh.getRange(2, col_(CFG.COL_DATA_MONTAGEM), lastRow - 1).setNumberFormat(
      "dd/MM/yyyy",
    );
  } catch (_) {}
  try {
    sh.getRange(2, col_(CFG.COL_DATA_TESTES), lastRow - 1).setNumberFormat(
      "dd/MM/yyyy",
    );
  } catch (_) {}
  try {
    sh.getRange(2, col_(CFG.COL_DATA_EVENTO), lastRow - 1).setNumberFormat(
      "dd/MM/yyyy",
    );
  } catch (_) {}
  try {
    sh.getRange(2, col_(CFG.COL_DATA_EMISSAO), lastRow - 1).setNumberFormat(
      "dd/MM/yyyy",
    );
  } catch (_) {}
  try {
    sh.getRange(2, col_(CFG.COL_VALIDADE), lastRow - 1).setNumberFormat(
      "dd/MM/yyyy",
    );
  } catch (_) {}

  try {
    sh.getRange(2, col_(CFG.COL_VALOR_ESTIMADO), lastRow - 1).setNumberFormat(
      "R$ #,##0.00",
    );
  } catch (_) {}

  try {
    sh.getRange(2, col_(CFG.COL_PROPOSAL_CODE), lastRow - 1).setNumberFormat(
      "0",
    );
  } catch (_) {}
}

function applyProtectionsAndHiding_() {
  const sh = getSheet_();
  try {
    const prot = sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    prot.forEach((p) => {
      if ((p.getDescription() || "").includes("[Sistema]")) p.remove();
    });
  } catch (_) {}

  const h = getHeaderMap_();
  const lastRow = Math.max(2, sh.getMaxRows());

  // Define quem pode editar colunas de sistema
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const owner = ss.getOwner() ? ss.getOwner().getEmail() : null;
  const allowedEditors = [];

  if (owner) allowedEditors.push(owner);
  if (Array.isArray(CFG.SYSTEM_ADMINS)) {
    CFG.SYSTEM_ADMINS.forEach((email) => {
      if (email && !allowedEditors.includes(email)) {
        allowedEditors.push(email);
      }
    });
  }

  // Cria proteções fortes para cada coluna de sistema
  CFG.SYSTEM_HEADERS.forEach((name) => {
    const col = h[name];
    if (!col) return;
    try {
      const rng = sh.getRange(2, col, lastRow - 1, 1);
      const p = rng.protect().setDescription("[Sistema] " + name);
      p.setWarningOnly(false); // bloqueia edição
      if (p.canDomainEdit()) {
        p.setDomainEdit(false); // impede "todos do domínio"
      }
      const editors = p.getEditors();
      if (editors && editors.length) {
        p.removeEditors(editors); // remove todos os outros
      }
      if (allowedEditors.length) {
        p.addEditors(allowedEditors); // só owner + admins
      }
    } catch (_) {}
  });

  // Proteção da linha de cabeçalho (row 1) — somente owner + SYSTEM_ADMINS podem renomear colunas
  try {
    const lastCol = sh.getMaxColumns();
    const headerRange = sh.getRange(1, 1, 1, lastCol);
    const pHeader = headerRange
      .protect()
      .setDescription("[Sistema] header_row");
    pHeader.setWarningOnly(false);
    if (pHeader.canDomainEdit()) pHeader.setDomainEdit(false);
    const hEditors = pHeader.getEditors();
    if (hEditors && hEditors.length) pHeader.removeEditors(hEditors);
    if (allowedEditors.length) pHeader.addEditors(allowedEditors);
  } catch (_) {}

  // Proteção da coluna proposta_aprovada (apenas editores autorizados)
  try {
    const colAprov = h[CFG.COL_PROPOSTA_APROVADA];
    if (colAprov) {
      const allowedAprov = [];
      if (owner) allowedAprov.push(owner);
      (CFG.PROPOSTA_APROVADA_EDITORS || []).forEach((email) => {
        if (email && !allowedAprov.includes(email)) allowedAprov.push(email);
      });
      const rngAprov = sh.getRange(2, colAprov, lastRow - 1, 1);
      const pAprov = rngAprov
        .protect()
        .setDescription("[Sistema] proposta_aprovada");
      pAprov.setWarningOnly(false);
      if (pAprov.canDomainEdit()) pAprov.setDomainEdit(false);
      const editors = pAprov.getEditors();
      if (editors && editors.length) pAprov.removeEditors(editors);
      if (allowedAprov.length) pAprov.addEditors(allowedAprov);
    }
  } catch (_) {}

  try {
    sh.showColumns(1, sh.getMaxColumns());
  } catch (_) {}
  CFG.HIDDEN_HEADERS.forEach((name) => {
    try {
      sh.hideColumns(col_(name));
    } catch (_) {}
  });
}

function applyZebraSheet_() {
  const sh = getSheet_();
  const lastRow = Math.max(2, sh.getMaxRows());
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return;

  const dataRows = lastRow - 1;
  const ODD = "#FFFFFF";
  const EVEN = "#F7F9FC";

  const zebraMatrix = Array.from({ length: dataRows }, (_, i) => {
    const rowNumber = i + 2;
    const color = rowNumber % 2 === 0 ? EVEN : ODD;
    return Array(lastCol).fill(color);
  });

  sh.getRange(2, 1, dataRows, lastCol).setBackgrounds(zebraMatrix);

  try {
    const statusCol = col_(CFG.COL_STATUS_USER);
    const statuses = sh.getRange(2, statusCol, dataRows, 1).getValues();

    const statusColors = statuses.map((r, i) => {
      const val = String(r[0] || "").trim();
      return [STATUS_BG_MAP[val] || ((i + 2) % 2 === 0 ? EVEN : ODD)];
    });

    sh.getRange(2, statusCol, dataRows, 1).setBackgrounds(statusColors);
  } catch (_) {}
}

function paintRowColors_(row) {
  const sh = getSheet_();
  const lastCol = sh.getLastColumn();
  const color = row % 2 === 0 ? "#F7F9FC" : "#FFFFFF";
  sh.getRange(row, 1, 1, lastCol).setBackground(color);
  try {
    const sCol = col_(CFG.COL_STATUS_USER);
    const val = String(sh.getRange(row, sCol).getValue() || "").trim();
    if (STATUS_BG_MAP[val])
      sh.getRange(row, sCol).setBackground(STATUS_BG_MAP[val]);
  } catch (_) {}
}

function applyConditionalFormatting_() {
  const sh = getSheet_();
  const lastRow = sh.getMaxRows();
  if (lastRow < 2) return;
  const rules = [];
  const addRule = (cols, bg) => {
    cols.forEach((n) => {
      try {
        const c = col_(n);
        rules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenBlank()
            .setBackground(bg)
            .setRanges([sh.getRange(2, c, lastRow - 1, 1)])
            .build(),
        );
      } catch (_) {}
    });
  };
  addRule(CFG.REQUIRED_GATING_HEADERS, "#F1F3F4");
  addRule(CFG.REQUIRED_USER_HEADERS, "#EFEFEF");
  sh.setConditionalFormatRules(rules);
}

function applyColumnWidths_() {
  const sh = getSheet_();
  const lastRow = Math.max(2, sh.getMaxRows());
  const lastCol = sh.getLastColumn();
  const W = { XS: 90, S: 130, M: 160, L: 200, XL: 260, XXL: 320, XXXL: 380 };
  const map = {
    [CFG.COL_CREATED_AT]: W.M,
    [CFG.COL_PROPOSAL_CODE]: W.M,
    [CFG.COL_PROPOSAL_VERSION]: W.S,
    [CFG.COL_STATUS_USER]: W.M,
    [CFG.COL_EMPRESA]: W.M,
    [CFG.COL_CONTATO_COMERCIAL]: W.M,
    [CFG.COL_ORIGEM]: W.M,
    [CFG.COL_CIDADE_LOCAL]: W.XL,
    [CFG.COL_DATA_MONTAGEM]: W.L,
    [CFG.COL_DATA_TESTES]: W.L,
    [CFG.COL_DATA_EVENTO]: W.L,
    [CFG.COL_CLIENTE]: W.L,
    [CFG.COL_CLIENTE_CONTATO]: W.L,
    [CFG.COL_CLIENTE_EMAIL]: W.XL,
    [CFG.COL_CLIENTE_TEL]: W.M,
    [CFG.COL_EVENTO_NOME]: W.XXL,
    [CFG.COL_VALOR_ESTIMADO]: W.M,
    [CFG.COL_FORMAS_PAGAMENTO]: W.XXL,
    [CFG.COL_DATA_EMISSAO]: W.M,
    [CFG.COL_VALIDADE]: W.M,
    [CFG.COL_RESP_CONTRATANTE]: W.L,
    [CFG.COL_FOLDER_LINK]: W.XXXL,
    [CFG.COL_DOC_LINK]: W.XXXL,
    [CFG.COL_PDF_LINK]: W.XXXL,
    [CFG.COL_TRELLO]: W.XXXL,
    [CFG.COL_ANO]: W.XS,
    [CFG.COL_OK]: W.XS,
    [CFG.COL_OBS]: W.XXXL,
  };
  Object.keys(map).forEach((k) => {
    try {
      const c = col_(k);
      if (!isColumnHidden_(sh, c)) sh.setColumnWidth(c, map[k]);
    } catch (_) {}
  });
  try {
    if (lastRow >= 2) {
      sh.getRange(2, 1, lastRow - 1, lastCol)
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle");
      const cCreated = col_(CFG.COL_CREATED_AT);
      sh.getRange(2, cCreated, lastRow - 1, 1).setNumberFormat(
        "dd/MM/yyyy HH:mm:ss",
      );
    }
  } catch (_) {}
  try {
    sh.getRange(1, 1, 1, lastCol)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(false)
      .setFontWeight("bold");
    sh.setRowHeight(1, 32);
  } catch (_) {}
}

function resetALL() {
  const sh = getSheet_();
  const totalRows = sh.getMaxRows();
  const totalCols = sh.getLastColumn();
  if (totalRows <= 1) {
    toast_("Nada para limpar.");
    return;
  }
  try {
    const protections =
      sh.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    protections.forEach((p) => {
      if ((p.getDescription() || "").includes("[Sistema]")) {
        try {
          p.remove();
        } catch (_) {}
      }
    });
  } catch (_) {}
  try {
    sh.getRange(2, 1, totalRows - 1, totalCols).clearContent();
    sh.getRange(2, 1, totalRows - 1, totalCols).setBackground("#FFFFFF");
  } catch (_) {}
  try {
    PropertiesService.getDocumentProperties().setProperty(
      CFG.PROP_KEY,
      String(CFG.PROP_START_AT),
    );
  } catch (_) {}
  try {
    applyProtectionsAndHiding_();
    applyZebraSheet_();
  } catch (_) {}
  toast_("Reset concluído.");
}

function hasProjectTrigger_(handlerName) {
  return ScriptApp.getProjectTriggers().some(
    (t) => t.getHandlerFunction && t.getHandlerFunction() === handlerName,
  );
}

function ensureRequiredTriggers_() {
  const ss = SpreadsheetApp.getActive();

  if (!hasProjectTrigger_("onEditInstallable")) {
    ScriptApp.newTrigger("onEditInstallable")
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }

  if (!hasProjectTrigger_("normalizeStatusColumnColors")) {
    ScriptApp.newTrigger("normalizeStatusColumnColors")
      .timeBased()
      .everyMinutes(1)
      .create();
  }

  if (!hasProjectTrigger_("processPendingNovoRows")) {
    ScriptApp.newTrigger("processPendingNovoRows")
      .timeBased()
      .everyMinutes(1)
      .create();
  }
}

function leads_setup() {
  applyValidations_();
  applyProtectionsAndHiding_();
  applyConditionalFormatting_();
  applyColumnWidths_();
  ensureNextCodeProperty_();
  backfillAno_();
  backfillComercialAll_();
  applyZebraSheet_();
  ensureRequiredTriggers_();
  SpreadsheetApp.flush();
  toast_("Setup Completo (v5.2).");
}

function FIX_TRIGGERS_NOW() {
  const ss = SpreadsheetApp.getActive();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach((t) => ScriptApp.deleteTrigger(t));

  // Gatilho de edição manual (UI)
  ScriptApp.newTrigger("onEditInstallable")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Gatilho de tempo: sincroniza cores da coluna STATUS a cada minuto.
  // Necessário porque o N8N atualiza o status via API externa e o onEdit
  // não dispara para alterações programáticas/externas.
  ScriptApp.newTrigger("normalizeStatusColumnColors")
    .timeBased()
    .everyMinutes(1)
    .create();

  ScriptApp.newTrigger("processPendingNovoRows")
    .timeBased()
    .everyMinutes(1)
    .create();

  toast_("Gatilhos Reiniciados.");
}
function reset() {
  resetALL();
}

function repaintAllRows() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("pipeline");

  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    paintRowColors_(row);
  }
}

/**
 * Normaliza as cores da coluna STATUS para todas as linhas já existentes,
 * sem mexer em nenhuma outra coluna.
 */
function normalizeStatusColumnColors() {
  const sh = getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const dataRows = lastRow - 1;
  const sCol = col_(CFG.COL_STATUS_USER);

  const ODD = "#FFFFFF";
  const EVEN = "#F7F9FC";

  const statuses = sh.getRange(2, sCol, dataRows, 1).getValues();
  const colors = statuses.map((r, i) => {
    const v = String(r[0] || "").trim();
    return [STATUS_BG_MAP[v] || ((i + 2) % 2 === 0 ? EVEN : ODD)];
  });

  sh.getRange(2, sCol, dataRows, 1).setBackgrounds(colors);
  try {
    toast_("STATUS: cores normalizadas (somente coluna STATUS).");
  } catch (_) {}
}

function resetApprovalWebhookForActiveRow() {
  const email = getActiveEditorEmail_();
  if (!canSetRestrictedStatus_(email) && !canEditPropostaAprovada_(email)) {
    toast_("Sem permissão para resetar o teste de Aprovado.");
    return;
  }

  const sh = getSheet_();
  const row = sh.getActiveCell().getRow();
  if (row < 2) {
    toast_("Selecione uma linha de dados para resetar o teste.");
    return;
  }

  const payload = buildN8NPayload_(row);
  const key = clearApprovalWebhookSent_(row, payload);
  releaseRowLock_(row);

  console.log(`Marcador de Aprovado removido: ${key}`);
  toast_(`Teste de Aprovado liberado para a linha ${row}.`);
}

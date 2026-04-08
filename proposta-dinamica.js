/** * PROPOSTA DINÂMICA v2 (API HUB)
 * Funções:
 * 1. BUILD: Cria proposta baseada em Origem/Empresa (Lógica original)
 * 2. SANITIZE: Remove preços de um documento existente (Nova lógica)
 */

const CONFIG = {
  TOKEN: "2FUyX994Z6dZ", // Token de segurança
  TEMPLATES_FOLDER_ID: "",
  DEFAULT_DEST_FOLDER_ID: "",
  DEFAULT_VALIDITY_DAYS: 15,
  COLLAPSE_ANO_PREFIX: true,
};

// Mapeamento de Templates (PRESERVADO)
const ORIGEM_TO_TEMPLATE_NAME = {
  Site: "Proposta template - em branco",
  Cliente: "Proposta template - em branco",
  Telefone: "Proposta template - em branco",
  WhatsApp: "Proposta template - em branco",
  Sesc: "Proposta template - Sesc",
  SESC: "Proposta template - Sesc",
  Mis: "Proposta template - MIS",
  MIS: "Proposta template - MIS",
  Integração: "Proposta template - em branco",
  "Teatro Santander": "Proposta template - Teatro Santander",
  "Teatro Iguatemi": "Proposta template - Teatro Iguatemi",
  "Teatro Multiplan": "Proposta template - Teatro Multiplan",
  "Teatro Villa Lobos": "Proposta template - Teatro Villa Lobos",
  "Rooftop 033": "Proposta template - Rooftop 033",
  "Teatro + Rooftop 033": "Proposta template - Teatro + Rooftop 033",
};

// Blocos de ORG (PRESERVADO)
const ORGS = {
  PLAY: {
    "org.razao_social": "Play Audiovisual Ltda.",
    "org.cnpj": "30.122.784/0001-91",
    "org.inscricao_estadual": "ISENTO",
    "org.endereco_linha":
      "Rua Jupi, 138 – Santo Amaro - São Paulo – SP – 04755-050",
    "org.telefone": "(11) 3803-8686",
    "org.site": "https://playaudiovisual.com.br",
  },
  PROJECOES: {
    "org.razao_social": "Play Projeções Ltda.",
    "org.cnpj": "11.110.413/0001-45",
    "org.inscricao_estadual": "ISENTO",
    "org.endereco_linha":
      "Rua Jupi, 138 – Santo Amaro - São Paulo – SP – 04755-050",
    "org.telefone": "(11) 3803-8686",
    "org.site": "https://playaudiovisual.com.br",
  },
  TOTAL: {
    "org.razao_social": "Play Total Ltda",
    "org.cnpj": "62.544.089/0001-04",
    "org.inscricao_estadual": "ISENTO",
    "org.endereco_linha":
      "Rua Jupi, 138 – Santo Amaro - São Paulo – SP – 04755-050",
    "org.telefone": "(11) 3803-8686",
    "org.site": "https://playaudiovisual.com.br",
  },
};

const ORG_ALIASES = {
  PLAY: [
    "play audiovisual",
    "play",
    "playaudiovisual",
    "play audiovisual ltda",
    "Play Audiovisual",
  ],
  PROJECOES: [
    "play projeções",
    "play projecoes",
    "projeções",
    "projecoes",
    "Play Projeções",
    "Play Projecoes",
  ],
  TOTAL: ["play total", "Play Total"],
};

/* =============== ROTEADOR CENTRAL (doPost) ================= */

function doPost(e) {
  try {
    // 1. Parse do Body
    const raw =
      e && e.postData && e.postData.contents ? e.postData.contents : "";
    let body = {};
    try {
      body = raw ? JSON.parse(raw) : {};
    } catch {
      body = parseFormUrlEncoded_(raw);
    }

    // 2. Segurança (Token)
    const token = body.token || (e && e.parameter && e.parameter.token) || "";
    if (String(token) !== CONFIG.TOKEN) {
      return respond({ ok: false, error: "unauthorized" });
    }

    // 3. Roteamento pela Ação
    const action = body.action || "build"; // Default é 'build' para manter compatibilidade

    if (action === "sanitize") {
      return executeSanitization(body);
    } else {
      return executeBuildProposal(body);
    }
  } catch (err) {
    return respond({ ok: false, error: String(err) });
  }
}

/* =============== FUNÇÃO 1: GERAR PROPOSTA (Lógica Original) ================= */

function executeBuildProposal(body) {
  const origem = String(body.origem || "").trim();
  const empresa = String(body.empresa || "").trim();
  const destFolderId =
    String(body.dest_folder_id || "").trim() || CONFIG.DEFAULT_DEST_FOLDER_ID;
  const templatesFolderId =
    String(body.templates_folder_id || "").trim() || CONFIG.TEMPLATES_FOLDER_ID;
  const docFileName = String(body.doc_file_name || "").trim();
  const ctx = body.context || {};

  if (!origem) return respond({ ok: false, error: "missing_origem" });
  if (!docFileName)
    return respond({ ok: false, error: "missing_doc_file_name" });
  if (!destFolderId)
    return respond({ ok: false, error: "missing_dest_folder_id" });

  const templateName =
    ORIGEM_TO_TEMPLATE_NAME[origem] || ORIGEM_TO_TEMPLATE_NAME["Site"];
  const templateId = getTemplateIdByName_(templateName, templatesFolderId);

  if (!templateId) {
    return respond({
      ok: false,
      error: "template_not_found",
      detail: { origem, templateName },
    });
  }

  // Copia template
  const destFolder = DriveApp.getFolderById(destFolderId);
  const templateFile = DriveApp.getFileById(templateId);
  const copy = templateFile.makeCopy(docFileName, destFolder);
  const docId = copy.getId();
  const docUrl = "https://docs.google.com/document/d/" + docId + "/edit";

  // Prepara dados
  const map = flatKV_(ctx);
  const orgKV = selectOrgByEmpresa_(empresa);
  Object.assign(map, orgKV);
  normalizeMap_(map);

  if (!map["proposta.numero"]) {
    const code = map.codigo_prop || map.proposal_code || "";
    const ver = String(map.versao_prop || "A").toUpperCase();
    map["proposta.numero"] = code ? code + "-" + ver : "0001-" + ver;
  }

  // Edição do Doc
  const doc = DocumentApp.openById(docId);
  const sections = [doc.getBody(), doc.getHeader(), doc.getFooter()];
  const keysInDoc = collectPlaceholdersInDoc_(sections);
  const mapToApply = filterMapByKeys_(map, keysInDoc);

  if (
    CONFIG.COLLAPSE_ANO_PREFIX &&
    keysInDoc.has("ano") &&
    keysInDoc.has("proposta.numero")
  ) {
    sections.forEach(
      (sec) => sec && collapseCompositePatterns_(sec, mapToApply),
    );
  }

  sections.forEach((sec) => sec && replaceEverywhereTargeted_(sec, mapToApply));
  doc.saveAndClose();

  return respond({
    ok: true,
    doc_proposta_id: docId,
    doc_url: docUrl,
    template_used: templateName,
  });
}

/* =============== FUNÇÃO 2: SANITIZAR PREÇOS (Nova Lógica) ================= */

function executeSanitization(body) {
  const docId = body.docId;
  if (!docId) return respond({ ok: false, error: "missing_docId" });

  try {
    const doc = DocumentApp.openById(docId);
    const bodySection = doc.getBody();

    // Regex para valores monetários (R$ ou numéricos soltos comuns em tabelas de preço)
    // Captura: R$ opcional, espaço opcional, números com ponto e vírgula
    const pattern = "(\\s?R\\$\\s?)?(\\d{1,3}(\\.\\d{3})*|\\d+),\\d{2}";

    // Substitui tudo por [--]
    bodySection.replaceText(pattern, "[--]");

    doc.saveAndClose();
    return respond({ ok: true, status: "sanitized", docId: docId });
  } catch (err) {
    return respond({
      ok: false,
      error: "sanitization_failed",
      detail: String(err),
    });
  }
}

/* =============== HELPERS (Preservados e Otimizados) =============== */

function getTemplateIdByName_(templateName, templatesFolderId) {
  if (templatesFolderId) {
    try {
      const folder = DriveApp.getFolderById(templatesFolderId);
      const files = folder.getFilesByName(templateName);
      if (files.hasNext()) return files.next().getId();
    } catch (_) {}
  }
  const it = DriveApp.getFilesByName(templateName);
  if (it.hasNext()) return it.next().getId();
  return "";
}

function flatKV_(obj) {
  const out = {};
  Object.keys(obj || {}).forEach((k) => (out[k] = obj[k]));
  return out;
}

function normalizeMap_(map) {
  // --- PONTE: planilha/workflow manda proposta_validade_data ---
  if (!map.validade && map.proposta_validade_data) {
    map.validade = map.proposta_validade_data;
  }

  if (map.proposta_validade_data && !map["proposta.validade_data"]) {
    map["proposta.validade_data"] = fmtDate_(map.proposta_validade_data);
  }

  // (mantém como está)
  if (map.proposta_data_emissao && !map["proposta.data_emissao"])
    map["proposta.data_emissao"] = fmtDate_(map.proposta_data_emissao);

  if (!map["proposta.validade_data"])
    map["proposta.validade_data"] = normalizeValidity_(
      map.validade,
      map["proposta.data_emissao"],
    );
  if (map.proposta_data_emissao && !map["proposta.data_emissao"])
    map["proposta.data_emissao"] = fmtDate_(map.proposta_data_emissao);
  if (!map["proposta.validade_data"])
    map["proposta.validade_data"] = normalizeValidity_(
      map.validade,
      map["proposta.data_emissao"],
    );

  // Mapeamentos de Cliente e Comercial (Resumidos para economizar espaço, mas lógica mantida)
  if (map.cliente) map["cliente.razao_social"] = map.cliente;
  if (map.cliente_contato_nome)
    map["cliente.contato_nome"] = map.cliente_contato_nome;
  if (map.cliente_email) map["cliente.contato_email"] = map.cliente_email;
  if (map.cliente_telefone)
    map["cliente.contato_telefone"] = map.cliente_telefone;
  if (map.contato_comercial)
    map["comercial.contato_nome"] = map.contato_comercial;
  if (map.comercial_email) map["comercial.email"] = map.comercial_email;
  if (map.comercial_telefone)
    map["comercial.telefone"] = map.comercial_telefone;

  if (map.evento_nome) map["evento.nome"] = map.evento_nome;
  if (map.cidade_local) map["evento.local"] = map.cidade_local;
  if (map.data_montagem)
    map["evento.montagem.data"] = fmtDate_(map.data_montagem);

  if (map.data_testes_ensaio) {
    // usa a data de testes/ensaio, se houver
    map["evento.teste.evento.data"] = fmtDate_(map.data_testes_ensaio);
  } else if (map.data_evento) {
    // fallback: se não tiver teste/ensaio, usa data do evento
    map["evento.teste.evento.data"] = fmtDate_(map.data_evento);
  }

  if (map.data_evento) map["evento.data"] = fmtDate_(map.data_evento);
}

function normalizeValidity_(validade, baseEmissaoBR) {
  if (!validade) {
    if (!baseEmissaoBR) return "";
    let d = parseDateBR_(baseEmissaoBR);
    d.setDate(d.getDate() + CONFIG.DEFAULT_VALIDITY_DAYS);
    return formatDateBR_(d);
  }
  const n = Number(String(validade).replace(",", "."));
  if (!isNaN(n) && n > 0) {
    let d = parseDateBR_(baseEmissaoBR || formatDateBR_(new Date()));
    d.setDate(d.getDate() + n);
    return formatDateBR_(d);
  }
  return fmtDate_(validade);
}

function selectOrgByEmpresa_(empresaRaw) {
  const s = String(empresaRaw || "")
    .replace(/\u00A0/g, " ")
    .trim()
    .toLowerCase();
  for (const code in ORG_ALIASES) {
    if (ORG_ALIASES[code].some((a) => s === a.toLowerCase())) return ORGS[code];
  }
  return ORGS[s.toUpperCase()] || ORGS.PLAY;
}

function collectPlaceholdersInDoc_(sections) {
  const set = new Set();
  sections.forEach((sec) => {
    if (sec) collectTokensFromContainer_(sec, set);
  });
  return set;
}

function collectTokensFromContainer_(container, set) {
  const n = container.getNumChildren();
  for (let i = 0; i < n; i++) {
    const ch = container.getChild(i);
    if (ch.getType() == DocumentApp.ElementType.TABLE) {
      const tbl = ch.asTable();
      for (let r = 0; r < tbl.getNumRows(); r++) {
        const row = tbl.getRow(r);
        for (let c = 0; c < row.getNumCells(); c++) {
          collectTokensFromContainer_(row.getCell(c), set);
        }
      }
    } else if (ch.editAsText) {
      const txt = ch.editAsText().getText() || "";
      const re = /\{\{\s*([^}]+?)\s*\}\}|\[\[\s*([^\]]+?)\s*\]\]/g;
      let m;
      while ((m = re.exec(txt))) {
        set.add((m[1] || m[2] || "").trim());
      }
    }
  }
}

function filterMapByKeys_(map, keysSet) {
  const out = {};
  Object.keys(map).forEach((k) => {
    if (keysSet.has(k)) out[k] = map[k];
  });
  return out;
}

function collapseCompositePatterns_(container, map) {
  const num = String(map["proposta.numero"] || "");
  if (!num) return;
  const composite =
    "\\{\\{\\s*ano\\s*\\}\\}\\s*[-–]\\s*\\{\\{\\s*proposta\\.numero\\s*\\}\\}";
  let m = null;
  while ((m = container.findText(composite, m))) {
    const t = m.getElement().asText();
    t.deleteText(m.getStartOffset(), m.getEndOffsetInclusive());
    t.insertText(m.getStartOffset(), num);
  }
}

function replaceEverywhereTargeted_(container, map) {
  const keys = Object.keys(map);
  if (!keys.length) return;
  keys.forEach((k) => {
    const v = map[k] == null ? "" : String(map[k]);
    const pat = "\\{\\{\\s*" + escapeReg_(k) + "\\s*\\}\\}";
    container.replaceText(pat, v);
  });
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
function parseFormUrlEncoded_(raw) {
  return {};
} // Simplificado para JSON apenas
function fmtDate_(v) {
  if (!v) return "";
  let s = String(v).trim();
  if (s.match(/^\d{4}-\d{2}-\d{2}/)) {
    // YYYY-MM-DD
    const [y, m, d] = s.split("T")[0].split("-");
    return `${d}/${m}/${y}`;
  }
  return s;
}
function parseDateBR_(str) {
  const [d, m, y] = String(str).split("/");
  return new Date(Number(y), Number(m) - 1, Number(d));
}
function formatDateBR_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}
function escapeReg_(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

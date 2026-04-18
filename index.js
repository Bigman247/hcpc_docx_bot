require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const Anthropic   = require('@anthropic-ai/sdk');
const axios       = require('axios');
const path        = require('path');
const fs          = require('fs');
const os          = require('os');
const pdfParse    = require('pdf-parse');
const nodemailer  = require('nodemailer');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType,
  ShadingType, HeadingLevel, PageBreak,
  PageOrientation, VerticalAlign,
} = require('docx');

// ─────────────────────────────────────────────────────────────
// INIT
// ─────────────────────────────────────────────────────────────
const bot       = new TelegramBot(process.env.TELEGRAM_BOT_TOKEN, { polling: true });
const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

const ADMIN_USERNAME  = '@Sni9pa';
const ADMIN_ID        = process.env.ADMIN_TELEGRAM_ID;
const PAYMENT_LINK    = 'https://flutterwave.com/pay/y1sdlmgf1io9';
const PRICE           = 'GBP 25';

const FREE_IDS = (process.env.FREE_USER_IDS || '')
  .split(',').map(s => s.trim()).filter(Boolean);

// ─────────────────────────────────────────────────────────────
// SESSIONS
// ─────────────────────────────────────────────────────────────
const sessions = {};

function getSession(chatId) {
  if (!sessions[chatId]) {
    sessions[chatId] = {
      profession:      null,
      documents:       [],
      paid:            false,
      userEmail:       null,
      userName:        null,
      awaitingEmail:   false,
      awaitingPayment: false,
      awaitingSopQ:    false,
    };
  }
  return sessions[chatId];
}

function clearSession(chatId) { sessions[chatId] = null; }

function isPaid(session, chatId) {
  return FREE_IDS.includes(String(chatId)) || session.paid;
}

// ─────────────────────────────────────────────────────────────
// PROFESSIONS
// ─────────────────────────────────────────────────────────────
const PROFESSIONS = [
  { code: 'arts',     label: 'Arts Therapists',                total: 14 },
  { code: 'bio',      label: 'Biomedical Scientists',           total: 15 },
  { code: 'pod',      label: 'Chiropodists / Podiatrists',      total: 15 },
  { code: 'clinical', label: 'Clinical Scientists',             total: 15 },
  { code: 'diet',     label: 'Dietitians',                      total: 15 },
  { code: 'hearing',  label: 'Hearing Aid Dispensers',          total: 15 },
  { code: 'ot',       label: 'Occupational Therapists',         total: 15 },
  { code: 'odp',      label: 'Operating Dept. Practitioners',   total: 15 },
  { code: 'ortho',    label: 'Orthoptists',                     total: 15 },
  { code: 'para',     label: 'Paramedics',                      total: 15 },
  { code: 'physio',   label: 'Physiotherapists',                total: 15 },
  { code: 'psych',    label: 'Practitioner Psychologists',      total: 15 },
  { code: 'prosth',   label: 'Prosthetists / Orthotists',       total: 15 },
  { code: 'radio',    label: 'Radiographers',                   total: 15 },
  { code: 'salt',     label: 'Speech and Language Therapists',  total: 15 },
];

function findProf(code) { return PROFESSIONS.find(p => p.code === code) || null; }

function profKeyboard() {
  const rows = [];
  for (let i = 0; i < PROFESSIONS.length; i += 2) {
    const row = [{ text: PROFESSIONS[i].label, callback_data: 'prof_' + PROFESSIONS[i].code }];
    if (PROFESSIONS[i + 1]) row.push({ text: PROFESSIONS[i + 1].label, callback_data: 'prof_' + PROFESSIONS[i + 1].code });
    rows.push(row);
  }
  return { inline_keyboard: rows };
}

// ─────────────────────────────────────────────────────────────
// FILE HELPERS
// ─────────────────────────────────────────────────────────────
const ALLOWED_EXTS = ['.pdf', '.jpg', '.jpeg', '.png', '.webp', '.gif'];
const ALLOWED_MIME = ['application/pdf', 'image/jpeg', 'image/png', 'image/webp', 'image/gif'];
const IMAGE_MIME   = { '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.png': 'image/png', '.webp': 'image/webp', '.gif': 'image/gif' };

async function downloadFile(fileId) {
  const info = await bot.getFile(fileId);
  const url  = 'https://api.telegram.org/file/bot' + process.env.TELEGRAM_BOT_TOKEN + '/' + info.file_path;
  const res  = await axios.get(url, { responseType: 'arraybuffer' });
  return { buffer: Buffer.from(res.data), fileName: path.basename(info.file_path) };
}

async function docsToBlocks(documents) {
  const maxChars = Math.floor(80000 / Math.max(documents.length, 1));
  const blocks   = [];
  for (const doc of documents) {
    const ext = path.extname(doc.name).toLowerCase();
    if (ext === '.pdf') {
      try {
        const data  = await pdfParse(doc.buffer);
        const text  = (data.text || '').slice(0, maxChars);
        const trunc = (data.text || '').length > maxChars;
        blocks.push({ type: 'text', text: '[PDF: ' + doc.name + (trunc ? ' truncated' : '') + ']\n\n' + text });
      } catch (e) {
        blocks.push({ type: 'text', text: '[PDF: ' + doc.name + ' error: ' + e.message + ']' });
      }
    } else if (IMAGE_MIME[ext]) {
      blocks.push({ type: 'image', source: { type: 'base64', media_type: IMAGE_MIME[ext], data: doc.buffer.toString('base64') } });
    }
  }
  return blocks;
}

// ─────────────────────────────────────────────────────────────
// CLAUDE PROMPTS
// ─────────────────────────────────────────────────────────────
function promptStandard(n, total, profLabel, docNames) {
  return 'You are an expert HCPC Standards of Proficiency mapping specialist.\n\n' +
    'PROFESSION: ' + profLabel + '\n' +
    'TASK: Map Standard ' + n + ' of ' + total + ' ONLY.\n\n' +
    'RULES:\n' +
    '1. Use the COMPLETE official HCPC SoP wording for ' + profLabel + ' Standard ' + n + ' and every sub-standard (' + n + '.1, ' + n + '.2 ... all of them, no exceptions).\n' +
    '2. Each sub-standard = one object in the array. Never group or skip any.\n' +
    '3. mappingType must be exactly one of: "Primary Qualification", "Professional Experience", "Partial Coverage", "Not Evidenced", "Insufficient Documentation"\n' +
    '4. Extract REAL course codes and module names from the documents for evidence. Never invent codes.\n' +
    '5. If not evidenced, evidence field = "Insufficient supporting documentation uploaded for this point"\n\n' +
    'RESPOND WITH VALID JSON ONLY. No markdown. No explanation. No code fences. Just the JSON object:\n' +
    '{\n' +
    '  "standardNumber": ' + n + ',\n' +
    '  "standardTitle": "exact official title for Standard ' + n + ' for ' + profLabel + '",\n' +
    '  "subStandards": [\n' +
    '    { "ref": "' + n + '.1", "requirement": "exact official wording", "mappingType": "Primary Qualification", "evidence": "COURSE_CODE - Module Name" },\n' +
    '    { "ref": "' + n + '.2", "requirement": "exact official wording", "mappingType": "Insufficient Documentation", "evidence": "Insufficient supporting documentation uploaded for this point" }\n' +
    '  ],\n' +
    '  "coverageSummary": "1-2 sentence summary",\n' +
    '  "covered": 3,\n' +
    '  "total": 5,\n' +
    '  "percent": 60\n' +
    '}\n\n' +
    'Documents uploaded: ' + docNames.join(', ');
}

function promptSummary(profLabel, results) {
  const data = results.map(r => ({
    number: r.standardNumber, title: r.standardTitle,
    covered: r.covered, total: r.total, percent: r.percent,
    gaps: r.error ? 'mapping error' :
      (r.subStandards || []).filter(s => s.mappingType === 'Not Evidenced' || s.mappingType === 'Insufficient Documentation').map(s => s.ref).join(', ') || 'none',
  }));
  return 'You are an HCPC mapping specialist for ' + profLabel + '.\n\n' +
    'Based on results below, write an overall summary.\n\n' +
    'DATA:\n' + JSON.stringify(data, null, 2) + '\n\n' +
    'RESPOND WITH VALID JSON ONLY. No markdown. No code fences:\n' +
    '{ "overallPercent": 75, "totalCovered": 45, "totalCount": 60, "strengths": "...", "gaps": "...", "recommendation": "..." }';
}

function promptSingleSop(question, profLabel, docNames) {
  return 'You are an HCPC SoP specialist for ' + profLabel + '.\n\n' +
    'Map this specific standard against the uploaded documents:\n"' + question + '"\n\n' +
    'RESPOND WITH VALID JSON ONLY. No markdown. No code fences:\n' +
    '{ "question": "' + question + '", "mappingType": "Primary Qualification", "evidence": ["COURSE_CODE - Module Name"], "assessment": "detailed paragraph", "recommendation": "Fully / Partially / Not evidenced", "missingEvidence": "what else would help" }\n\n' +
    'Documents: ' + docNames.join(', ');
}

// ─────────────────────────────────────────────────────────────
// JSON PARSER
// ─────────────────────────────────────────────────────────────
function safeJSON(text) {
  if (!text) return null;
  const clean = text.replace(/```json|```/gi, '').trim();
  try { return JSON.parse(clean); } catch {}
  const m = clean.match(/\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch {} }
  return null;
}

// ─────────────────────────────────────────────────────────────
// DOCX BUILDER
// ─────────────────────────────────────────────────────────────
// A4 Landscape — content width after 600 DXA margins each side = 16838 - 1200 = 15638
const CW   = 15638;
const COLS = [900, 5000, 2838, 6900]; // Ref | Requirement | Type | Evidence — sum = 15638

const C = {
  darkBlue: '1F4E79', midBlue: '2E75B6',
  green: '1A6B1A', amber: 'B86B00', red: 'CC0000',
  altRow: 'EBF3FF', white: 'FFFFFF',
};

function borders(color) {
  const b = { style: BorderStyle.SINGLE, size: 4, color };
  return { top: b, bottom: b, left: b, right: b };
}

function tc(text, w, opts) {
  const o = opts || {};
  return new TableCell({
    borders:       borders(o.bc || 'CCCCCC'),
    width:         { size: w, type: WidthType.DXA },
    shading:       { fill: o.bg || C.white, type: ShadingType.CLEAR },
    margins:       { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.TOP,
    children: [new Paragraph({ children: [new TextRun({
      text:    String(text || ''),
      bold:    o.bold || false,
      italics: o.italic || false,
      color:   o.color || '111111',
      size:    o.size || 17,
      font:    'Arial',
    })] })],
  });
}

function hc(text, w) {
  return tc(text, w, { bold: true, color: 'FFFFFF', bg: C.darkBlue, bc: C.darkBlue, size: 18 });
}

function typeColor(t) {
  if (!t) return '555555';
  const l = t.toLowerCase();
  if (l.includes('primary') || l.includes('professional')) return C.green;
  if (l.includes('partial')) return C.amber;
  return C.red;
}

function sp() { return new Paragraph({ children: [new TextRun({ text: '' })] }); }

function h2(text) {
  return new Paragraph({
    spacing: { before: 200, after: 80 },
    border:  { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.midBlue, space: 1 } },
    children: [new TextRun({ text, bold: true, size: 26, color: C.darkBlue, font: 'Arial' })],
  });
}

function kv(label, value, lc) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: label, bold: true, size: 18, color: lc || C.darkBlue, font: 'Arial' }),
      new TextRun({ text: String(value || ''), size: 18, color: '222222', font: 'Arial' }),
    ],
  });
}

async function buildDocx(profLabel, results, summary) {
  const kids = [];

  // Cover
  kids.push(
    new Paragraph({ spacing: { before: 1400, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'HCPC Standards of Proficiency', bold: true, size: 56, color: C.darkBlue, font: 'Arial' })] }),
    new Paragraph({ spacing: { before: 60, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Mapping Report', bold: true, size: 38, color: C.midBlue, font: 'Arial' })] }),
    new Paragraph({ spacing: { before: 0, after: 60 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: profLabel, bold: true, size: 30, color: '444444', font: 'Arial' })] }),
    new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.midBlue, space: 1 } }, children: [new TextRun({ text: 'Generated: ' + new Date().toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }), size: 20, color: '666666', font: 'Arial' })] }),
    sp(), sp(),
  );

  // Legend
  kids.push(h2('Mapping Key'));
  const legendItems = [
    ['Primary Qualification',      C.green, 'Directly and fully evidenced in the curriculum or module content'],
    ['Professional Experience',    C.green, 'Evidenced through clinical placement or practice components'],
    ['Partial Coverage',           C.amber, 'Partially addressed — supplementary evidence may be needed'],
    ['Not Evidenced',              C.red,   'Not found in the uploaded documents'],
    ['Insufficient Documentation', C.red,   'Documents do not contain enough detail to map this point'],
  ];
  const lw = [3000, 12638];
  kids.push(
    new Table({
      width: { size: CW, type: WidthType.DXA }, columnWidths: lw,
      rows: [
        new TableRow({ tableHeader: true, children: [hc('Mapping Type', lw[0]), hc('Meaning', lw[1])] }),
        ...legendItems.map(([type, color, meaning], i) => new TableRow({ children: [
          tc(type, lw[0], { bold: true, color, bg: i % 2 === 0 ? C.altRow : C.white }),
          tc(meaning, lw[1], { bg: i % 2 === 0 ? C.altRow : C.white }),
        ]})),
      ],
    }),
    sp(),
    new Paragraph({ children: [new PageBreak()] }),
  );

  // Per-standard tables
  for (const std of results) {
    kids.push(h2('Standard ' + std.standardNumber + ': ' + (std.standardTitle || 'Standard ' + std.standardNumber)));

    if (std.error) {
      kids.push(new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: 'Could not map this standard: ' + std.error, color: C.red, size: 18, italics: true, font: 'Arial' })] }));
    } else {
      const rows = std.subStandards || [];
      kids.push(new Table({
        width: { size: CW, type: WidthType.DXA }, columnWidths: COLS,
        rows: [
          new TableRow({ tableHeader: true, children: [hc('Ref', COLS[0]), hc('HCPC Requirement', COLS[1]), hc('Mapping Type', COLS[2]), hc('Evidence (Course Code / Module)', COLS[3])] }),
          ...rows.map((row, i) => {
            const bg      = i % 2 === 0 ? C.altRow : C.white;
            const tc_color = typeColor(row.mappingType);
            const insuff  = (row.evidence || '').toLowerCase().includes('insufficient');
            return new TableRow({ children: [
              tc(row.ref,         COLS[0], { bold: true, color: C.darkBlue, bg }),
              tc(row.requirement, COLS[1], { bg }),
              tc(row.mappingType, COLS[2], { bold: true, color: tc_color, bg }),
              tc(row.evidence,    COLS[3], { italic: insuff, color: insuff ? C.red : '222222', bg }),
            ]});
          }),
        ],
      }));
    }

    kids.push(
      sp(),
      kv('Coverage Summary: ', std.coverageSummary || ''),
      kv('Sub-standards covered: ', (std.covered || '0') + ' of ' + (std.total || '0') + '  |  Coverage: ' + (std.percent || '0') + '%', C.midBlue),
      sp(),
      new Paragraph({ children: [new PageBreak()] }),
    );
  }

  // Overall summary table
  kids.push(h2('Overall Summary'));
  const sw = [700, 4000, 2200, 1500, 7238];
  kids.push(
    new Table({
      width: { size: CW, type: WidthType.DXA }, columnWidths: sw,
      rows: [
        new TableRow({ tableHeader: true, children: [hc('Std', sw[0]), hc('Standard Title', sw[1]), hc('Covered', sw[2]), hc('Coverage', sw[3]), hc('Gaps / Notes', sw[4])] }),
        ...results.map((std, i) => {
          const bg  = i % 2 === 0 ? C.altRow : C.white;
          const pct = std.percent || 0;
          const sc  = pct >= 80 ? C.green : pct >= 50 ? C.amber : C.red;
          const gaps = std.error ? 'Mapping error' :
            (std.subStandards || []).filter(s => s.mappingType === 'Not Evidenced' || s.mappingType === 'Insufficient Documentation').map(s => s.ref).join(', ') || 'None';
          return new TableRow({ children: [
            tc(String(std.standardNumber), sw[0], { bold: true, color: C.darkBlue, bg }),
            tc(std.standardTitle || 'Standard ' + std.standardNumber, sw[1], { bg }),
            tc((std.covered || '0') + ' / ' + (std.total || '0'), sw[2], { bg }),
            tc(pct + '%', sw[3], { bold: true, color: sc, bg }),
            tc(gaps, sw[4], { italic: gaps !== 'None', color: gaps === 'None' ? C.green : '444444', bg }),
          ]});
        }),
      ],
    }),
    sp(),
  );

  if (summary) {
    kids.push(
      kv('Overall Coverage: ', (summary.overallPercent || '0') + '%  (' + (summary.totalCovered || '0') + ' of ' + (summary.totalCount || '0') + ' sub-standards)'),
      sp(),
      kv('Strengths: ', summary.strengths || '', C.green),
      sp(),
      kv('Gaps: ', summary.gaps || '', C.red),
      sp(),
      kv('Recommendation: ', summary.recommendation || '', C.darkBlue),
    );
  }

  const doc = new Document({
    styles: {
      default: { document: { run: { font: 'Arial', size: 20 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 56, bold: true, font: 'Arial', color: C.darkBlue }, paragraph: { spacing: { before: 480, after: 240 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { size: 26, bold: true, font: 'Arial', color: C.darkBlue }, paragraph: { spacing: { before: 200, after: 120 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE }, margin: { top: 600, right: 600, bottom: 600, left: 600 } } },
      children: kids,
    }],
  });

  return Packer.toBuffer(doc);
}

// ─────────────────────────────────────────────────────────────
// EMAIL
// ─────────────────────────────────────────────────────────────
async function sendEmail(toEmail, profLabel, docxBuffer, userName) {
  const transport = nodemailer.createTransport({
    host: process.env.SMTP_HOST, port: parseInt(process.env.SMTP_PORT || '587'),
    secure: false, auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
  });
  const safeName = profLabel.replace(/[^a-zA-Z0-9]/g, '_');
  const fileName = 'HCPC_SOP_' + safeName + '_' + Date.now() + '.docx';
  await transport.sendMail({
    from:    '"HCPC SOP Mapper" <' + process.env.SMTP_USER + '>',
    to:      toEmail,
    subject: 'Your HCPC SOP Mapping Report - ' + profLabel,
    html: '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #dde4f0;border-radius:10px;overflow:hidden;">' +
      '<div style="background:#1F4E79;padding:24px 28px;"><h1 style="color:#fff;margin:0;font-size:20px;">HCPC SOP Mapping Report</h1><p style="color:#93C6E7;margin:4px 0 0;font-size:14px;">' + profLabel + '</p></div>' +
      '<div style="background:#f7faff;padding:24px 28px;"><p>Dear ' + (userName || 'Healthcare Professional') + ',</p>' +
      '<p>Your HCPC Standards of Proficiency mapping report is attached as a Word document.</p>' +
      '<p><strong>The report includes:</strong></p><ul><li>Every standard and sub-standard mapped individually</li><li>Exact HCPC requirement wording per sub-standard</li><li>Evidence from your uploaded course documents</li><li>Clear flags where documentation is insufficient</li><li>Overall summary with coverage percentages</li></ul>' +
      '<p style="margin-top:20px;padding:14px;background:#EBF3FF;border-radius:6px;font-size:13px;">Open the .docx file in Microsoft Word or Google Docs for best results.</p>' +
      '<hr style="border:none;border-top:1px solid #dde4f0;margin:20px 0;">' +
      '<p style="color:#888;font-size:12px;margin:0;">Generated by HCPC SOP Mapper | Need help? Contact <a href="https://t.me/Sni9pa" style="color:#1F4E79;">@Sni9pa</a></p></div></div>',
    attachments: [{ filename: fileName, content: docxBuffer }],
  });
  return fileName;
}

// ─────────────────────────────────────────────────────────────
// CORE MAPPING PIPELINE
// ─────────────────────────────────────────────────────────────
async function runMapping(chatId, session) {
  const prof      = findProf(session.profession);
  const profLabel = prof.label;
  const total     = prof.total;
  const docNames  = session.documents.map(d => d.name);

  const hdr = await bot.sendMessage(chatId,
    '*HCPC SOP Mapping Started*\n\n' +
    'Profession: ' + profLabel + '\n' +
    'Documents: ' + session.documents.length + '\n' +
    'Standards to map: ' + total + '\n' +
    'Report email: ' + session.userEmail + '\n\n' +
    'Reading your documents now...',
    { parse_mode: 'Markdown' }
  );

  // Build content blocks
  let blocks;
  try {
    blocks = await docsToBlocks(session.documents);
  } catch (err) {
    return bot.editMessageText('Failed to read documents: ' + err.message + '\n\nContact ' + ADMIN_USERNAME, { chat_id: chatId, message_id: hdr.message_id });
  }

  await bot.editMessageText(
    '*Mapping ' + total + ' standards one by one...*\n\nProfession: ' + profLabel + '\n\nProgress updates appear below as each standard completes:',
    { chat_id: chatId, message_id: hdr.message_id, parse_mode: 'Markdown' }
  );

  const results = [];

  for (let n = 1; n <= total; n++) {
    const tick = await bot.sendMessage(chatId, 'Mapping Standard ' + n + ' of ' + total + '...');
    try {
      const resp = await anthropic.messages.create({
        model: 'claude-opus-4-5', max_tokens: 8192,
        messages: [{ role: 'user', content: [...blocks, { type: 'text', text: promptStandard(n, total, profLabel, docNames) }] }],
      });
      const raw    = resp.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
      const parsed = safeJSON(raw);
      if (!parsed || !Array.isArray(parsed.subStandards)) throw new Error('AI returned invalid format');
      results.push(parsed);
      const pct   = parsed.percent || 0;
      const emoji = pct >= 80 ? 'done' : pct >= 50 ? 'partial' : 'low';
      await bot.editMessageText(
        'Standard ' + n + '/' + total + ' mapped: ' + (parsed.standardTitle || '') + '\n' +
        'Sub-standards: ' + (parsed.covered || 0) + '/' + (parsed.total || 0) + ' | Coverage: ' + pct + '%',
        { chat_id: chatId, message_id: tick.message_id }
      );
    } catch (err) {
      console.error('Standard ' + n + ' error:', err.message);
      results.push({ standardNumber: n, standardTitle: 'Standard ' + n, error: err.message, subStandards: [], coverageSummary: 'Could not map: ' + err.message, covered: 0, total: 0, percent: 0 });
      await bot.editMessageText('Standard ' + n + '/' + total + ': Could not map - ' + err.message, { chat_id: chatId, message_id: tick.message_id }).catch(() => {});
    }
    await new Promise(r => setTimeout(r, 350));
  }

  // Summary
  const sumTick = await bot.sendMessage(chatId, 'Generating overall summary...');
  let summary = null;
  try {
    const sumResp = await anthropic.messages.create({
      model: 'claude-opus-4-5', max_tokens: 2048,
      messages: [{ role: 'user', content: [{ type: 'text', text: promptSummary(profLabel, results) }] }],
    });
    summary = safeJSON(sumResp.content.filter(b => b.type === 'text').map(b => b.text).join('\n'));
  } catch (err) { console.error('Summary error:', err.message); }

  await bot.editMessageText('Building Word document...', { chat_id: chatId, message_id: sumTick.message_id });

  // Build docx
  let docxBuffer;
  try {
    docxBuffer = await buildDocx(profLabel, results, summary);
  } catch (err) {
    console.error('DOCX error:', err);
    return bot.editMessageText('Failed to build document: ' + err.message + '\n\nContact ' + ADMIN_USERNAME, { chat_id: chatId, message_id: sumTick.message_id });
  }

  await bot.editMessageText('Sending report to your email...', { chat_id: chatId, message_id: sumTick.message_id });

  let sent = false;
  try {
    const fn = await sendEmail(session.userEmail, profLabel, docxBuffer, session.userName);
    await bot.editMessageText(
      '*Report emailed successfully!*\n\nSent to: ' + session.userEmail + '\nFile: ' + fn + '\n\nCheck your inbox (and spam folder if not seen within 5 minutes).',
      { chat_id: chatId, message_id: sumTick.message_id, parse_mode: 'Markdown' }
    );
    sent = true;
  } catch (emailErr) {
    console.error('Email error:', emailErr.message);
    await bot.editMessageText('Email failed - sending via Telegram instead...', { chat_id: chatId, message_id: sumTick.message_id });
  }

  if (!sent) {
    try {
      const tmp = path.join(os.tmpdir(), 'hcpc_' + chatId + '_' + Date.now() + '.docx');
      fs.writeFileSync(tmp, docxBuffer);
      await bot.sendDocument(chatId, tmp, {}, { filename: 'HCPC_SOP_' + profLabel.replace(/[^a-zA-Z0-9]/g, '_') + '.docx', contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      fs.unlinkSync(tmp);
    } catch (tgErr) {
      await bot.sendMessage(chatId, 'Could not deliver report. Contact ' + ADMIN_USERNAME);
    }
  }

  const mapped = results.filter(r => !r.error).length;
  const failed = results.filter(r => !!r.error).length;
  await bot.sendMessage(chatId,
    '*Mapping Complete!*\n\n' +
    'Profession: ' + profLabel + '\n' +
    mapped + '/' + total + ' standards mapped' +
    (failed > 0 ? '\n' + failed + ' had errors (noted in report)' : '') +
    (summary ? '\nOverall coverage: ' + (summary.overallPercent || 0) + '%' : '') + '\n\n' +
    '/mapsop - map a specific standard in detail\n' +
    '/reset - start a new report\n' +
    'Contact ' + ADMIN_USERNAME + ' for questions',
    { parse_mode: 'Markdown' }
  );
}

// ─────────────────────────────────────────────────────────────
// BOT COMMANDS
// ─────────────────────────────────────────────────────────────
bot.onText(/\/start/, msg => {
  const chatId = msg.chat.id;
  getSession(chatId);
  bot.sendMessage(chatId,
    '*HCPC SOP Mapper*\n\nMaps your course documents to the exact HCPC Standards of Proficiency for your profession, then emails you a complete Word document report.\n\nSelect your profession to begin:',
    { parse_mode: 'Markdown', reply_markup: profKeyboard() }
  );
});

bot.onText(/\/help/, msg => {
  bot.sendMessage(msg.chat.id,
    '*HCPC SOP Mapper - Help*\n\n' +
    'Accepted files: PDF, JPG, PNG, WEBP\n' +
    'Output: Professional Word document emailed to you\n\n' +
    'Steps:\n' +
    '1. /start - select profession\n' +
    '2. /upload - send your course documents\n' +
    '3. /pay - pay ' + PRICE + '\n' +
    '4. /generate - enter email and start mapping\n' +
    '5. Receive full report by email\n\n' +
    'Other commands:\n' +
    '/mapsop - map one specific SOP question\n' +
    '/status - check your session\n' +
    '/reset - start over\n' +
    '/profession - change profession\n\n' +
    'Contact ' + ADMIN_USERNAME + ' for support',
    { parse_mode: 'Markdown' }
  );
});

bot.onText(/\/status/, msg => {
  const chatId  = msg.chat.id;
  const session = getSession(chatId);
  const prof    = findProf(session.profession);
  const isFree  = FREE_IDS.includes(String(chatId));
  const docs    = session.documents.length > 0 ? session.documents.map((d, i) => (i + 1) + '. ' + d.name).join('\n') : 'None uploaded yet';
  bot.sendMessage(chatId,
    '*Your Session*\n\n' +
    'Profession: ' + (prof ? prof.label : 'Not selected') + '\n' +
    'Documents:\n' + docs + '\n' +
    'Email: ' + (session.userEmail || 'Not set') + '\n' +
    'Payment: ' + (isFree ? 'Admin/Free Access' : session.paid ? 'Confirmed' : 'Not confirmed - use /pay') + '\n\n' +
    'Contact ' + ADMIN_USERNAME + ' for help',
    { parse_mode: 'Markdown' }
  );
});

bot.onText(/\/reset/, msg => {
  clearSession(msg.chat.id);
  bot.sendMessage(msg.chat.id, 'Session cleared. Select your profession to start again:', { reply_markup: profKeyboard() });
});

bot.onText(/\/profession/, msg => {
  bot.sendMessage(msg.chat.id, 'Select your HCPC profession:', { reply_markup: profKeyboard() });
});

bot.onText(/\/upload/, msg => {
  const chatId  = msg.chat.id;
  const session = getSession(chatId);
  if (!session.profession) return bot.sendMessage(chatId, 'Select your profession first:', { reply_markup: profKeyboard() });
  bot.sendMessage(chatId, 'Send your documents now.\n\nAccepted: PDF, JPG, PNG, WEBP\nNot accepted: Word - convert to PDF first\n\nSend files then use /generate when ready.');
});

bot.onText(/\/pay/, msg => {
  const chatId  = msg.chat.id;
  const session = getSession(chatId);
  if (FREE_IDS.includes(String(chatId))) return bot.sendMessage(chatId, 'You have free access - use /generate!');
  session.awaitingPayment = true;
  bot.sendMessage(chatId,
    '*Payment Instructions*\n\nPay here: ' + PAYMENT_LINK + '\nAmount: ' + PRICE + '\n\nAfter paying:\n1. Screenshot your payment confirmation\n2. Send the screenshot here\n3. We verify and unlock /generate\n\nContact ' + ADMIN_USERNAME + ' for help',
    { parse_mode: 'Markdown' }
  );
});

bot.onText(/\/confirm_payment (.+)/, (msg, match) => {
  if (String(msg.chat.id) !== String(ADMIN_ID)) return bot.sendMessage(msg.chat.id, 'Unauthorised.');
  const targetId = match[1].trim();
  const s = getSession(targetId);
  s.paid = true;
  bot.sendMessage(targetId, 'Payment confirmed! Use /generate to start your mapping.');
  bot.sendMessage(msg.chat.id, 'Payment confirmed for ' + targetId);
});

bot.onText(/\/generate/, async msg => {
  const chatId  = msg.chat.id;
  const session = getSession(chatId);
  if (!session.profession) return bot.sendMessage(chatId, 'Select profession first:', { reply_markup: profKeyboard() });
  if (session.documents.length === 0) return bot.sendMessage(chatId, 'Upload documents first with /upload');
  if (!isPaid(session, chatId)) return bot.sendMessage(chatId, 'Payment required. Use /pay.\n\n' + PAYMENT_LINK + '\nAmount: ' + PRICE);
  if (!session.userEmail) {
    session.awaitingEmail = true;
    return bot.sendMessage(chatId, 'Enter your email address.\n\nYour full HCPC SOP mapping report will be emailed as a Word document.\n\nType your email now:');
  }
  await runMapping(chatId, session);
});

bot.onText(/\/mapsop/, msg => {
  const chatId  = msg.chat.id;
  const session = getSession(chatId);
  if (!session.profession) return bot.sendMessage(chatId, 'Select profession first:', { reply_markup: profKeyboard() });
  if (session.documents.length === 0) return bot.sendMessage(chatId, 'Upload documents first with /upload');
  if (!isPaid(session, chatId)) return bot.sendMessage(chatId, 'Payment required. Use /pay first.');
  session.awaitingSopQ = true;
  bot.sendMessage(chatId,
    '*Single SOP Question Mapper*\n\nType the specific HCPC standard or sub-standard you want to map.\n\nExamples:\n"2.1 - maintain high standards of personal conduct"\n"13.5 - understand radiation protection"\n\nType your question now:',
    { parse_mode: 'Markdown' }
  );
});

// ─────────────────────────────────────────────────────────────
// FILE HANDLERS
// ─────────────────────────────────────────────────────────────
bot.on('document', async msg => {
  const chatId = msg.chat.id;
  const doc    = msg.document;
  const ext    = path.extname(doc.file_name || '').toLowerCase();
  if (!ALLOWED_MIME.includes(doc.mime_type) && !ALLOWED_EXTS.includes(ext)) {
    return bot.sendMessage(chatId, 'File type not supported. Accepted: PDF, JPG, PNG, WEBP. Convert Word docs to PDF first. Contact ' + ADMIN_USERNAME);
  }
  const tick = await bot.sendMessage(chatId, 'Receiving ' + doc.file_name + '...');
  try {
    const { buffer, fileName } = await downloadFile(doc.file_id);
    const session = getSession(chatId);
    session.documents.push({ name: doc.file_name || fileName, buffer });
    await bot.editMessageText(doc.file_name + ' received! Total: ' + session.documents.length + ' document(s). Send more or use /generate when ready.', { chat_id: chatId, message_id: tick.message_id });
  } catch (err) {
    await bot.editMessageText('Failed to receive file: ' + err.message, { chat_id: chatId, message_id: tick.message_id });
  }
});

bot.on('photo', async msg => {
  const chatId = msg.chat.id;
  const photo  = msg.photo[msg.photo.length - 1];
  const tick   = await bot.sendMessage(chatId, 'Receiving image...');
  try {
    const { buffer } = await downloadFile(photo.file_id);
    const session    = getSession(chatId);
    session.documents.push({ name: 'image_' + Date.now() + '.jpg', buffer });
    await bot.editMessageText('Image received! Total: ' + session.documents.length + ' document(s). Send more or use /generate when ready.', { chat_id: chatId, message_id: tick.message_id });
  } catch (err) {
    await bot.editMessageText('Failed to receive image: ' + err.message, { chat_id: chatId, message_id: tick.message_id });
  }
});

// ─────────────────────────────────────────────────────────────
// CALLBACK QUERIES
// ─────────────────────────────────────────────────────────────
bot.on('callback_query', async query => {
  const chatId  = query.message.chat.id;
  const data    = query.data;
  const session = getSession(chatId);
  await bot.answerCallbackQuery(query.id);
  if (data.startsWith('prof_')) {
    const code = data.replace('prof_', '');
    const prof = findProf(code);
    if (!prof) return;
    session.profession = code;
    await bot.editMessageText('Profession selected: ' + prof.label + '\nStandards to map: ' + prof.total, { chat_id: chatId, message_id: query.message.message_id });
    bot.sendMessage(chatId, 'Next steps:\n\n1. /upload - send your course documents\n2. /pay - pay ' + PRICE + '\n3. /generate - enter email and start mapping\n\nYour report will be emailed as a Word document.\n\n/help for more info | ' + ADMIN_USERNAME + ' for support');
  }
});

// ─────────────────────────────────────────────────────────────
// MESSAGE CATCH-ALL
// ─────────────────────────────────────────────────────────────
bot.on('message', async msg => {
  const chatId  = msg.chat.id;
  const session = getSession(chatId);
  const text    = msg.text;
  if (!text || text.startsWith('/') || msg.document || msg.photo) return;

  // Email collection
  if (session.awaitingEmail) {
    const emailRx = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRx.test(text.trim())) return bot.sendMessage(chatId, 'That does not look like a valid email address. Please try again:');
    session.userEmail     = text.trim();
    session.awaitingEmail = false;
    session.userName      = msg.from && msg.from.first_name ? msg.from.first_name : '';
    await bot.sendMessage(chatId, 'Got it! Report will be sent to ' + session.userEmail + '\n\nStarting your full mapping now...');
    await runMapping(chatId, session);
    return;
  }

  // Single SOP question
  if (session.awaitingSopQ) {
    session.awaitingSopQ = false;
    const question = text.trim();
    const prof     = findProf(session.profession);
    const tick     = await bot.sendMessage(chatId, 'Mapping: "' + question + '"\n\nPlease wait...');
    try {
      const blocks = await docsToBlocks(session.documents);
      blocks.push({ type: 'text', text: promptSingleSop(question, prof.label, session.documents.map(d => d.name)) });
      const resp   = await anthropic.messages.create({ model: 'claude-opus-4-5', max_tokens: 4096, messages: [{ role: 'user', content: blocks }] });
      const raw    = resp.content.filter(b => b.type === 'text').map(b => b.text).join('\n');
      const parsed = safeJSON(raw);
      await bot.deleteMessage(chatId, tick.message_id).catch(() => {});
      if (parsed) {
        const ev = (parsed.evidence || []).map(e => '- ' + e).join('\n') || 'None found';
        await bot.sendMessage(chatId,
          '*Single SOP Mapping Result*\n\n' +
          'Standard: ' + (parsed.question || question) + '\n\n' +
          'Mapping Type: ' + (parsed.mappingType || '-') + '\n\n' +
          'Evidence found:\n' + ev + '\n\n' +
          'Assessment: ' + (parsed.assessment || '-') + '\n\n' +
          'Recommendation: ' + (parsed.recommendation || '-') +
          (parsed.missingEvidence ? '\n\nMissing evidence: ' + parsed.missingEvidence : ''),
          { parse_mode: 'Markdown' }
        );
      } else {
        await bot.sendMessage(chatId, raw.slice(0, 4000));
      }
      await bot.sendMessage(chatId, 'Done.\n/mapsop - map another\n/generate - full report');
    } catch (err) {
      await bot.editMessageText('Error: ' + err.message + '\n\nContact ' + ADMIN_USERNAME, { chat_id: chatId, message_id: tick.message_id }).catch(() => {});
    }
    return;
  }

  // Payment note
  if (session.awaitingPayment) {
    bot.sendMessage(chatId, 'Thanks. Once payment is verified you will get a confirmation.\n\n/status to check | ' + ADMIN_USERNAME + ' if urgent');
  }
});

// ─────────────────────────────────────────────────────────────
console.log('HCPC SOP Docx Mapper Bot is running');

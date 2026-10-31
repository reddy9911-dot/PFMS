/**
 * PFMS - Factura auto-fill.
 *
 * Costa Rica electronic invoices (XML) are read in the browser for free.
 * This file adds the optional AI reader for PDFs and photos, which sends the
 * file to the Claude API and returns the same fields.
 *
 * The AI reader is off until an Anthropic API key is saved on the Settings
 * screen (stored as the PFMS_ANTHROPIC_API_KEY Script Property, never sent
 * back to the browser). Each document read is billed to that key.
 */

var PFMS_ANTHROPIC_KEY_PROPERTY = 'PFMS_ANTHROPIC_API_KEY';
var PFMS_CLAUDE_MODEL = 'claude-opus-5-5';
var PFMS_CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
var PFMS_CLAUDE_API_VERSION = '2026-10-21';
var PFMS_AI_MAX_BYTES = 20 * 1024 * 1024;
var PFMS_AI_MEDIA_TYPES = ['application/pdf', 'image/jpeg', 'image/png', 'image/gif', 'image/webp'];

var PFMS_FACTURA_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['is_invoice', 'document_type', 'total', 'currency', 'date', 'supplier_name', 'supplier_id',
    'customer_name', 'invoice_number', 'description', 'suggested_category'],
  properties: {
    is_invoice: { type: 'boolean' },
    document_type: { type: 'string', enum: ['invoice', 'receipt', 'credit_note', 'debit_note', 'other'] },
    total: { anyOf: [{ type: 'number' }, { type: 'null' }] },
    currency: { type: 'string' },
    date: { type: 'string' },
    supplier_name: { type: 'string' },
    supplier_id: { type: 'string' },
    customer_name: { type: 'string' },
    invoice_number: { type: 'string' },
    description: { type: 'string' },
    suggested_category: { type: 'string' }
  }
};

/**
 * @return {{aiEnabled: boolean}}
 */
function getFacturaReaderStatus() {
  return { aiEnabled: !!getAnthropicApiKey_() };
}

/**
 * Checks the key with Anthropic (a free models request) before saving it.
 * @param {string} apiKey
 * @return {{aiEnabled: boolean}}
 */
function saveAnthropicApiKey(apiKey) {
  var key = String(apiKey || '').trim();
  if (!/^sk-ant-[A-Za-z0-9_\-]{20,}$/.test(key)) {
    throw new Error('That doesn\'t look like an Anthropic API key (it should start with sk-ant-).');
  }

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/models?limit=1', {
    method: 'get',
    headers: { 'x-api-key': key, 'anthropic-version': PFMS_CLAUDE_API_VERSION },
    muteHttpExceptions: true
  });
  var status = response.getResponseCode();
  if (status === 401 || status === 403) {
    throw new Error('Anthropic rejected that key. Check it and try again.');
  }
  if (status >= 400) {
    throw new Error('Could not check the key with Anthropic right now (HTTP ' + status + '). Try again later.');
  }

  PropertiesService.getScriptProperties().setProperty(PFMS_ANTHROPIC_KEY_PROPERTY, key);
  return { aiEnabled: true };
}

/**
 * @return {{aiEnabled: boolean}}
 */
function removeAnthropicApiKey() {
  PropertiesService.getScriptProperties().deleteProperty(PFMS_ANTHROPIC_KEY_PROPERTY);
  return { aiEnabled: false };
}

/**
 * Reads a PDF or photo of an invoice with Claude.
 *
 * @param {{name: string, mimeType: string, base64: string}} upload
 * @param {{paymentType: string, categories: string[]}} context
 * @return {Object} the fields found (same shape as the browser's XML reader)
 */
function readFacturaWithAI(upload, context) {
  var apiKey = getAnthropicApiKey_();
  if (!apiKey) {
    throw new Error('AI reading is off. Add an Anthropic API key in Settings.');
  }
  if (!upload || !upload.base64) {
    throw new Error('No file was received.');
  }

  var mediaType = String(upload.mimeType || '').toLowerCase();
  if (PFMS_AI_MEDIA_TYPES.indexOf(mediaType) === -1) {
    throw new Error('AI reading works with PDF, JPEG, PNG, GIF or WebP files.');
  }
  if (upload.base64.length * 0.75 > PFMS_AI_MAX_BYTES) {
    throw new Error('The file is too large to read (20 MB maximum).');
  }

  context = context || {};
  var isIncome = String(context.paymentType || '').toLowerCase() === 'income';
  var categories = (context.categories || []).map(function (c) {
    return String(c || '').trim();
  }).filter(Boolean).slice(0, 50);

  var fileBlock = mediaType === 'application/pdf'
    ? { type: 'document', source: { type: 'base64', media_type: mediaType, data: upload.base64 } }
    : { type: 'image', source: { type: 'base64', media_type: mediaType, data: upload.base64 } };

  var instructions = [
    'This file was attached to a ' + (isIncome ? 'payment received' : 'payment made') +
      ' in a household property-management app in Costa Rica. Read the invoice, receipt or factura and fill in the fields.',
    'total: the final amount payable (total comprobante / total a pagar), as a plain number; null if you cannot find it.',
    'currency: ISO code such as CRC or USD (₡ or "colones" means CRC); empty string if unclear.',
    'date: the issue date as YYYY-MM-DD; empty string if unclear. Costa Rican dates are usually day/month/year.',
    'supplier_name / supplier_id: who issued the document and their cédula or tax ID (digits only).',
    'customer_name: who it was issued to.',
    'invoice_number: the consecutive or invoice number.',
    'description: a few words saying what was bought or charged.',
    categories.length
      ? 'suggested_category: the best match from this list, copied exactly, or empty string if none fits: ' + categories.join(' | ')
      : 'suggested_category: empty string.',
    'is_invoice: false if the file is not an invoice or receipt at all.',
    'Use empty strings for anything not shown. Do not guess values that are not in the document.'
  ].join('\n');

  var body = {
    model: PFMS_CLAUDE_MODEL,
    max_tokens: 16000,
    fallbacks: 'default',
    output_config: {
      effort: 'low',
      format: { type: 'json_schema', schema: PFMS_FACTURA_SCHEMA }
    },
    messages: [{
      role: 'user',
      content: [fileBlock, { type: 'text', text: instructions }]
    }]
  };

  var response = UrlFetchApp.fetch(PFMS_CLAUDE_API_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': PFMS_CLAUDE_API_VERSION,
      'anthropic-beta': 'server-side-fallback-2026-07-01'
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  var status = response.getResponseCode();
  var text = response.getContentText();
  if (status !== 200) {
    throw new Error(describeClaudeError_(status, text));
  }

  return parseClaudeFacturaResponse_(JSON.parse(text), categories);
}

/* ---------- helpers ---------- */

function getAnthropicApiKey_() {
  return PropertiesService.getScriptProperties().getProperty(PFMS_ANTHROPIC_KEY_PROPERTY) || '';
}

function describeClaudeError_(status, text) {
  var message = '';
  try {
    message = (JSON.parse(text).error || {}).message || '';
  } catch (e) {
    // Not JSON; use the status alone.
  }
  if (status === 401 || status === 403) {
    return 'The Anthropic API key was rejected. Update it in Settings.';
  }
  if (status === 429) {
    return 'The AI reader is busy or out of credit. Try again in a minute, or check your Anthropic account.';
  }
  if (status === 413) {
    return 'The file is too large for the AI reader.';
  }
  if (status >= 500) {
    return 'The AI reader is temporarily unavailable. Try again shortly.';
  }
  return 'The AI reader could not read this file' + (message ? ': ' + message : '.');
}

function parseClaudeFacturaResponse_(message, categories) {
  if (message.stop_reason === 'refusal') {
    throw new Error('The AI reader declined to read this file.');
  }
  if (message.stop_reason === 'max_tokens') {
    throw new Error('The AI reader\'s answer was cut off. Try again.');
  }

  var textBlock = (message.content || []).filter(function (block) {
    return block.type === 'text' && block.text;
  })[0];
  if (!textBlock) {
    throw new Error('The AI reader returned no answer.');
  }

  var data = JSON.parse(textBlock.text);
  var currency = String(data.currency || '').toUpperCase().trim();
  var date = /^\d{4}-\d{2}-\d{2}$/.test(data.date || '') ? data.date : '';
  var category = String(data.suggested_category || '').trim();
  if (categories.indexOf(category) === -1) {
    category = '';
  }

  return {
    source: 'ai',
    isInvoice: data.is_invoice !== false,
    documentType: data.document_type || 'other',
    total: typeof data.total === 'number' && isFinite(data.total) ? Math.abs(data.total) : null,
    currency: currency,
    date: date,
    supplierName: String(data.supplier_name || '').trim(),
    supplierId: String(data.supplier_id || '').replace(/\D/g, ''),
    customerName: String(data.customer_name || '').trim(),
    invoiceNumber: String(data.invoice_number || '').trim(),
    description: String(data.description || '').trim(),
    suggestedCategory: category
  };
}

/**
 * Lovely Finds / Cantinho dos Achados Publisher
 * Google Apps Script backend.
 *
 * Secrets are stored in Script Properties, never in HTML or GitHub:
 * META_IG_ACCESS_TOKEN
 * IG_USER_ID
 * META_API_VERSION (optional; defaults to v27.0)
 */

const APP_CONFIG = {
  spreadsheetProperty: 'PUBLISHER_SPREADSHEET_ID',
  tokenProperty: 'META_IG_ACCESS_TOKEN',
  igUserIdProperty: 'IG_USER_ID',
  apiVersionProperty: 'META_API_VERSION',
  defaultApiVersion: 'v27.0',
  postsSheet: 'Posts'
};

const POST_HEADERS = [
  'id', 'createdAt', 'productName', 'imageUrl', 'caption',
  'status', 'containerId', 'mediaId', 'publishedAt', 'error'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Cantinho dos Achados Publisher')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboardData() {
  const spreadsheet = getOrCreateSpreadsheet_();
  const posts = readPosts_(spreadsheet).slice(-30).reverse();
  const profile = testInstagramConnection_();
  return {
    brand: 'Cantinho dos Achados',
    profile: profile,
    posts: posts,
    storageUrl: spreadsheet.getUrl(),
    apiVersion: getApiVersion_()
  };
}

function testInstagramConnection() {
  return testInstagramConnection_();
}

function saveDraft(payload) {
  validateDraft_(payload);
  const spreadsheet = getOrCreateSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(APP_CONFIG.postsSheet);
  const id = Utilities.getUuid();
  const row = [
    id,
    new Date().toISOString(),
    String(payload.productName || 'Publicação sem título').trim().slice(0, 120),
    String(payload.imageUrl).trim(),
    String(payload.caption || '').trim().slice(0, 2200),
    'DRAFT',
    '',
    '',
    '',
    ''
  ];
  sheet.appendRow(row);
  return { ok: true, id: id, message: 'Rascunho salvo.' };
}

function publishDraft(draftId) {
  if (!draftId || typeof draftId !== 'string') {
    throw new Error('Identificador do rascunho inválido.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('Outra publicação está em andamento. Aguarde alguns segundos.');
  }

  try {
    const token = getRequiredProperty_(APP_CONFIG.tokenProperty);
    const igUserId = getRequiredProperty_(APP_CONFIG.igUserIdProperty);
    const spreadsheet = getOrCreateSpreadsheet_();
    const sheet = spreadsheet.getSheetByName(APP_CONFIG.postsSheet);
    const found = findPostRow_(sheet, draftId);
    if (!found) throw new Error('Rascunho não encontrado.');

    const post = rowToPost_(found.values);
    if (post.status === 'PUBLISHED') throw new Error('Este rascunho já foi publicado.');

    updatePostCell_(sheet, found.row, 6, 'PROCESSING');
    updatePostCell_(sheet, found.row, 10, '');

    try {
      // Reuse an existing container after a temporary processing timeout.
      let containerId = post.containerId;
      if (!containerId || post.status === 'DRAFT' || post.status === 'ERROR') {
        const createResponse = metaRequest_(
          'POST',
          '/' + encodeURIComponent(igUserId) + '/media',
          { image_url: post.imageUrl, caption: post.caption },
          token
        );
        containerId = createResponse.id;
        if (!containerId) throw new Error('A Meta não retornou o ID do contêiner.');
        updatePostCell_(sheet, found.row, 7, containerId);
      }

      let status = '';
      for (let attempt = 0; attempt < 25; attempt++) {
        const statusResponse = metaRequest_(
          'GET',
          '/' + encodeURIComponent(containerId),
          { fields: 'status_code' },
          token
        );
        status = statusResponse.status_code || '';
        if (status === 'FINISHED') break;
        if (status === 'ERROR' || status === 'EXPIRED') {
          throw new Error('A Meta informou que o contêiner ficou ' + status + '.');
        }
        Utilities.sleep(1500);
      }

      if (status !== 'FINISHED') {
        updatePostCell_(sheet, found.row, 6, 'PROCESSING');
        return {
          ok: true,
          pending: true,
          containerId: containerId,
          message: 'A imagem ainda está sendo processada. Clique em “Retomar publicação” daqui a pouco.'
        };
      }

      const publishResponse = metaRequest_(
        'POST',
        '/' + encodeURIComponent(igUserId) + '/media_publish',
        { creation_id: containerId },
        token
      );
      if (!publishResponse.id) throw new Error('A Meta não retornou o ID da publicação.');

      updatePostCell_(sheet, found.row, 6, 'PUBLISHED');
      updatePostCell_(sheet, found.row, 8, publishResponse.id);
      updatePostCell_(sheet, found.row, 9, new Date().toISOString());

      return {
        ok: true,
        pending: false,
        mediaId: publishResponse.id,
        containerId: containerId,
        message: 'Publicação enviada ao Instagram com sucesso.'
      };
    } catch (error) {
      updatePostCell_(sheet, found.row, 6, 'ERROR');
      updatePostCell_(sheet, found.row, 10, safeErrorMessage_(error));
      throw new Error(safeErrorMessage_(error));
    }
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const existingId = properties.getProperty(APP_CONFIG.spreadsheetProperty);
  let spreadsheet;
  if (existingId) {
    try {
      spreadsheet = SpreadsheetApp.openById(existingId);
    } catch (error) {
      properties.deleteProperty(APP_CONFIG.spreadsheetProperty);
    }
  }
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create('Cantinho dos Achados — Publicações');
    properties.setProperty(APP_CONFIG.spreadsheetProperty, spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName(APP_CONFIG.postsSheet);
  if (!sheet) sheet = spreadsheet.insertSheet(APP_CONFIG.postsSheet);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, POST_HEADERS.length).setValues([POST_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, POST_HEADERS.length).setFontWeight('bold');
  }
  const defaultSheet = spreadsheet.getSheetByName('Sheet1');
  if (defaultSheet && spreadsheet.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    spreadsheet.deleteSheet(defaultSheet);
  }
  return spreadsheet;
}

function readPosts_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(APP_CONFIG.postsSheet);
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, POST_HEADERS.length)
    .getValues()
    .map(rowToPost_);
}

function rowToPost_(row) {
  return {
    id: String(row[0] || ''),
    createdAt: String(row[1] || ''),
    productName: String(row[2] || ''),
    imageUrl: String(row[3] || ''),
    caption: String(row[4] || ''),
    status: String(row[5] || 'DRAFT'),
    containerId: String(row[6] || ''),
    mediaId: String(row[7] || ''),
    publishedAt: String(row[8] || ''),
    error: String(row[9] || '')
  };
}

function findPostRow_(sheet, id) {
  if (sheet.getLastRow() < 2) return null;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, POST_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === id) return { row: i + 2, values: values[i] };
  }
  return null;
}

function updatePostCell_(sheet, row, column, value) {
  sheet.getRange(row, column).setValue(value);
}

function validateDraft_(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Dados do rascunho inválidos.');
  const imageUrl = String(payload.imageUrl || '').trim();
  if (!/^https:\/\//i.test(imageUrl)) throw new Error('A imagem precisa ter um endereço HTTPS público.');
  if (imageUrl.length > 2000) throw new Error('O endereço da imagem é muito longo.');
  const caption = String(payload.caption || '').trim();
  if (caption.length > 2200) throw new Error('A legenda excede 2.200 caracteres.');
}

function testInstagramConnection_() {
  const token = getRequiredProperty_(APP_CONFIG.tokenProperty);
  const igUserId = getRequiredProperty_(APP_CONFIG.igUserIdProperty);
  const result = metaRequest_(
    'GET',
    '/' + encodeURIComponent(igUserId),
    { fields: 'id,username' },
    token
  );
  return {
    connected: true,
    id: result.id || '',
    username: result.username || '',
    matchesConfiguredAccount: String(result.id) === String(igUserId)
  };
}

function metaRequest_(method, path, params, token) {
  const apiVersion = getApiVersion_();
  const baseUrl = 'https://graph.facebook.com/' + apiVersion + path;
  const options = {
    method: method.toLowerCase(),
    muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + token }
  };
  let url = baseUrl;
  if (method === 'GET') {
    const query = Object.keys(params || {}).map(function(key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(String(params[key]));
    }).join('&');
    if (query) url += '?' + query;
  } else {
    options.payload = params || {};
  }

  const response = UrlFetchApp.fetch(url, options);
  const text = response.getContentText();
  let data;
  try { data = JSON.parse(text); } catch (error) {
    throw new Error('A Meta retornou uma resposta que não pôde ser interpretada.');
  }
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300 || data.error) {
    const apiError = data.error || {};
    throw new Error(
      'Meta API (' + (apiError.code || response.getResponseCode()) + '): ' +
      (apiError.message || 'A solicitação não foi concluída.')
    );
  }
  return data;
}

function getRequiredProperty_(name) {
  const value = PropertiesService.getScriptProperties().getProperty(name);
  if (!value) throw new Error('Configuração ausente no Apps Script: ' + name + '.');
  return value;
}

function getApiVersion_() {
  return PropertiesService.getScriptProperties().getProperty(APP_CONFIG.apiVersionProperty) ||
    APP_CONFIG.defaultApiVersion;
}

function safeErrorMessage_(error) {
  const message = String(error && error.message ? error.message : error || 'Erro desconhecido');
  return message.replace(/EA[A-Za-z0-9_-]{30,}/g, '[TOKEN REDIGIDO]').slice(0, 600);
}

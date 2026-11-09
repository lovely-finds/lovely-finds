/**
 * Lovely Finds / Cantinho dos Achados Publisher
 * Google Apps Script backend.
 *
 * Secrets are stored in Script Properties, never in HTML or GitHub:
 * META_IG_ACCESS_TOKEN
 * IG_USER_ID
 * META_API_VERSION (optional; defaults to v27.0)
 * GITHUB_TOKEN (fine-grained token with Contents: Read and write)
 */

const APP_CONFIG = {
  spreadsheetProperty: 'PUBLISHER_SPREADSHEET_ID',
  tokenProperty: 'META_IG_ACCESS_TOKEN',
  igUserIdProperty: 'IG_USER_ID',
  apiVersionProperty: 'META_API_VERSION',
  githubTokenProperty: 'GITHUB_TOKEN',
  githubRepository: 'lovely-finds/lovely-finds',
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

      // Mark the post as published BEFORE cleanup. A GitHub cleanup failure
      // must never turn a successful Instagram publication into an ERROR.
      updatePostCell_(sheet, found.row, 6, 'PUBLISHED');
      updatePostCell_(sheet, found.row, 8, publishResponse.id);
      updatePostCell_(sheet, found.row, 9, new Date().toISOString());

      let cleanup = { ok: false, skipped: true, message: 'Limpeza não executada.' };
      try {
        cleanup = deletePublishedImageFromGitHub_(post.imageUrl);
        updatePostCell_(sheet, found.row, 10, cleanup.ok ? '' : cleanup.message);
      } catch (cleanupError) {
        cleanup = { ok: false, message: safeErrorMessage_(cleanupError) };
        updatePostCell_(sheet, found.row, 10, 'Publicado; falha ao remover imagem do GitHub: ' + cleanup.message);
      }

      return {
        ok: true,
        pending: false,
        mediaId: publishResponse.id,
        containerId: containerId,
        cleanup: cleanup,
        message: cleanup.ok
          ? 'Publicação enviada ao Instagram e imagem removida da pasta uploads.'
          : 'Publicação enviada ao Instagram, mas a imagem permaneceu no GitHub para limpeza manual.'
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


/**
 * Safe image URL test: asks Meta to create and process a media container.
 * IMPORTANT: this function never calls /media_publish and cannot publish a post.
 * Run manually from the Apps Script editor when validating a public image URL.
 */
function testMetaImageUrl() {
  const imageUrl = 'https://raw.githubusercontent.com/lovely-finds/lovely-finds/main/uploads/teste-upload-cantinho.jpg';
  if (!/^https:\/\//i.test(imageUrl)) throw new Error('A URL de teste precisa ser HTTPS.');

  const token = getRequiredProperty_(APP_CONFIG.tokenProperty);
  const igUserId = getRequiredProperty_(APP_CONFIG.igUserIdProperty);

  const created = metaRequest_(
    'POST',
    '/' + encodeURIComponent(igUserId) + '/media',
    {
      image_url: imageUrl,
      caption: 'Teste técnico de imagem — Cantinho dos Achados'
    },
    token
  );

  if (!created.id) throw new Error('A Meta não retornou o ID do contêiner.');

  let status = '';
  let details = null;
  for (let attempt = 0; attempt < 20; attempt++) {
    details = metaRequest_(
      'GET',
      '/' + encodeURIComponent(created.id),
      { fields: 'status_code,status' },
      token
    );
    status = details.status_code || details.status || '';
    if (status === 'FINISHED' || status === 'ERROR' || status === 'EXPIRED') break;
    Utilities.sleep(1500);
  }

  const result = {
    ok: status === 'FINISHED',
    published: false,
    containerId: created.id,
    status: status || 'IN_PROGRESS',
    message: status === 'FINISHED'
      ? 'A Meta conseguiu buscar e processar a imagem. Nenhuma publicação foi feita.'
      : 'O contêiner foi criado, mas o processamento ainda não foi confirmado. Nenhuma publicação foi feita.'
  };
  Logger.log(JSON.stringify(result));
  return result;
}


/**
 * Removes only image files inside uploads/ from the configured public repository.
 * The GitHub commit history is preserved; this only removes the file from current main.
 */
function deletePublishedImageFromGitHub_(imageUrl) {
  const rawPrefix = 'https://raw.githubusercontent.com/lovely-finds/lovely-finds/main/';
  const url = String(imageUrl || '').trim();
  if (url.indexOf(rawPrefix) !== 0) {
    return { ok: false, skipped: true, message: 'URL fora do repositório autorizado; arquivo não removido.' };
  }

  const path = decodeURIComponent(url.substring(rawPrefix.length).split('?')[0]);
  if (!/^uploads\/[A-Za-z0-9._-]+\.(jpg|jpeg|png|webp)$/i.test(path)) {
    return { ok: false, skipped: true, message: 'Caminho fora de uploads/ ou extensão não permitida; arquivo não removido.' };
  }

  const token = getRequiredProperty_(APP_CONFIG.githubTokenProperty);
  const apiUrl = 'https://api.github.com/repos/' + APP_CONFIG.githubRepository +
    '/contents/' + path.split('/').map(encodeURIComponent).join('/');
  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };

  const getResponse = UrlFetchApp.fetch(apiUrl + '?ref=main', {
    method: 'get',
    headers: headers,
    muteHttpExceptions: true
  });
  const getCode = getResponse.getResponseCode();
  if (getCode === 404) {
    return { ok: true, alreadyAbsent: true, path: path, message: 'A imagem já não estava na pasta uploads/.' };
  }
  if (getCode < 200 || getCode >= 300) {
    throw new Error('GitHub não conseguiu localizar a imagem (HTTP ' + getCode + ').');
  }

  const file = JSON.parse(getResponse.getContentText());
  if (!file.sha) throw new Error('GitHub não retornou o SHA do arquivo para exclusão.');

  const deleteResponse = UrlFetchApp.fetch(apiUrl, {
    method: 'delete',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify({
      message: 'Remove published image: ' + path,
      sha: file.sha,
      branch: 'main'
    }),
    muteHttpExceptions: true
  });
  const deleteCode = deleteResponse.getResponseCode();
  if (deleteCode < 200 || deleteCode >= 300) {
    throw new Error('GitHub não conseguiu remover a imagem (HTTP ' + deleteCode + ').');
  }

  return {
    ok: true,
    path: path,
    commit: (JSON.parse(deleteResponse.getContentText()).commit || {}).sha || '',
    message: 'Imagem removida da pasta uploads/; histórico do Git preservado.'
  };
}

/**
 * Manual diagnostic: checks GitHub token permissions without deleting anything.
 */
function testGitHubCleanupConnection() {
  const token = getRequiredProperty_(APP_CONFIG.githubTokenProperty);
  const response = UrlFetchApp.fetch(
    'https://api.github.com/repos/' + APP_CONFIG.githubRepository,
    {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token,
        Accept: 'application/vnd.github+json',
        'X-GitHub-Api-Version': '2022-11-28'
      },
      muteHttpExceptions: true
    }
  );
  const status = response.getResponseCode();
  if (status < 200 || status >= 300) {
    throw new Error('Conexão GitHub falhou (HTTP ' + status + '). Verifique o token e as permissões.');
  }
  const data = JSON.parse(response.getContentText());
  return { ok: true, repository: data.full_name, message: 'Token aceito pelo GitHub.' };
}


/**
 * End-to-end permission test. Creates a uniquely named tiny PNG in uploads/
 * and immediately removes it using the same cleanup helper used after publishing.
 * It does not call Instagram or publish anything.
 */
function testGitHubUploadAndDelete() {
  const token = getRequiredProperty_(APP_CONFIG.githubTokenProperty);
  const path = 'uploads/test-cleanup-permission.png';
  const imageBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/pWQAAAAASUVORK5CYII=';
  const apiUrl = 'https://api.github.com/repos/' + APP_CONFIG.githubRepository +
    '/contents/' + path.split('/').map(encodeURIComponent).join('/');
  const headers = {
    Authorization: 'Bearer ' + token,
    Accept: 'application/vnd.github+json',
    'X-GitHub-Api-Version': '2022-11-28'
  };

  // Never overwrite a pre-existing file.
  const existing = UrlFetchApp.fetch(apiUrl + '?ref=main', {
    method: 'get', headers: headers, muteHttpExceptions: true
  });
  if (existing.getResponseCode() === 200) {
    throw new Error('O arquivo de teste já existe. Não foi alterado nem removido.');
  }
  if (existing.getResponseCode() !== 404) {
    throw new Error('Não foi possível verificar o arquivo de teste (HTTP ' + existing.getResponseCode() + ').');
  }

  const create = UrlFetchApp.fetch(apiUrl, {
    method: 'put',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify({
      message: 'Test GitHub upload permission for Apps Script cleanup',
      content: imageBase64,
      branch: 'main'
    }),
    muteHttpExceptions: true
  });
  if (create.getResponseCode() < 200 || create.getResponseCode() >= 300) {
    throw new Error('GitHub não conseguiu criar o arquivo de teste (HTTP ' + create.getResponseCode() + ').');
  }

  // Delete only after GitHub confirms the test file was created.
  const deletion = deletePublishedImageFromGitHub_(
    'https://raw.githubusercontent.com/' + APP_CONFIG.githubRepository + '/main/' + path
  );
  if (!deletion.ok) {
    throw new Error('O arquivo de teste foi criado, mas não foi removido: ' + deletion.message);
  }

  return {
    ok: true,
    created: true,
    deleted: true,
    path: path,
    historyPreserved: true,
    message: 'O Apps Script criou e removeu o arquivo de teste com sucesso. Nenhuma publicação foi feita.'
  };
}

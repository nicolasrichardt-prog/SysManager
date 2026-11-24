const crypto = require('crypto');
const webpush = require('web-push');

const DEFAULT_VAPID_PUBLIC_KEY = 'BJqDopb3x4JAaKo4evVkZ6Ksq2l37E6uc83GjqXC1g7v-u0WnNhItnKJF331nVOMEis9SC0AZipnlgnne2d4rzk';

function setCorsHeaders(res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Push-Secret');
}

function timingSafeCompare(a, b) {
  if (!a || !b) return false;
  const hashA = crypto.createHash('sha256').update(String(a)).digest();
  const hashB = crypto.createHash('sha256').update(String(b)).digest();
  return crypto.timingSafeEqual(hashA, hashB);
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    if (req.body !== undefined && req.body !== null) {
      if (typeof req.body === 'object' && !Buffer.isBuffer(req.body)) {
        resolve(req.body);
        return;
      }
      try {
        const str = Buffer.isBuffer(req.body) ? req.body.toString('utf8') : String(req.body);
        resolve(str ? JSON.parse(str) : {});
        return;
      } catch (err) {
        reject(new Error('JSON inválido no corpo da requisição.'));
        return;
      }
    }

    const chunks = [];
    req.on('data', chunk => chunks.push(chunk));
    req.on('end', () => {
      try {
        const raw = Buffer.concat(chunks).toString('utf8').trim();
        resolve(raw ? JSON.parse(raw) : {});
      } catch (err) {
        reject(new Error('JSON inválido no stream da requisição.'));
      }
    });
    req.on('error', reject);
  });
}

module.exports = async function handler(req, res) {
  setCorsHeaders(res);

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  const publicKey = (process.env.VAPID_PUBLIC_KEY || DEFAULT_VAPID_PUBLIC_KEY || '').trim();
  const privateKey = (process.env.VAPID_PRIVATE_KEY || '').trim();
  const secret = (process.env.PUSH_API_SECRET || '').trim();
  const subject = (process.env.VAPID_SUBJECT || 'mailto:sysmanager.sm@gmail.com').trim();

  if (req.method === 'GET') {
    const hasPublicKey = Boolean(publicKey);
    const hasPrivateKey = Boolean(privateKey);
    const hasSecret = Boolean(secret);
    return res.status(200).json({
      ok: true,
      configured: hasPublicKey && hasPrivateKey && hasSecret,
      hasPublicKey: hasPublicKey,
      hasPrivateKey: hasPrivateKey,
      hasSecret: hasSecret,
      publicKey: publicKey
    });
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, error: 'Método não permitido. Use GET ou POST.' });
  }

  const missingVars = [];
  if (!publicKey) missingVars.push('VAPID_PUBLIC_KEY');
  if (!privateKey) missingVars.push('VAPID_PRIVATE_KEY');
  if (!secret) missingVars.push('PUSH_API_SECRET');

  if (missingVars.length > 0) {
    return res.status(500).json({
      success: false,
      error: 'Variáveis de ambiente ausentes na Vercel: ' + missingVars.join(', '),
      missingVars: missingVars
    });
  }

  let body;
  try {
    body = await readJsonBody(req);
  } catch (errParse) {
    return res.status(400).json({ success: false, error: errParse.message });
  }

  const authHeader = String(req.headers.authorization || '').trim();
  const bearerToken = authHeader.toLowerCase().startsWith('bearer ') ? authHeader.slice(7).trim() : '';
  const providedSecret = String(
    req.headers['x-push-secret'] || bearerToken || (body && body.secret) || ''
  ).trim();

  if (!timingSafeCompare(providedSecret, secret)) {
    return res.status(401).json({ success: false, error: 'Segredo PUSH_API_SECRET inválido ou ausente.' });
  }

  try {
    webpush.setVapidDetails(subject, publicKey, privateKey);
  } catch (errVapid) {
    return res.status(500).json({
      success: false,
      error: 'Erro ao configurar chaves VAPID: ' + (errVapid.message || String(errVapid))
    });
  }

  const items = Array.isArray(body.items)
    ? body.items
    : (body.subscription ? [{
        id: body.id || '',
        endpoint: body.subscription.endpoint || '',
        subscription: body.subscription,
        payload: body.payload || {},
        ttl: body.ttl
      }] : []);

  if (items.length === 0) {
    return res.status(400).json({
      success: false,
      error: 'Nenhum item ou subscription informado para envio.'
    });
  }

  let sent = 0;
  let failed = 0;
  const expiredEndpoints = [];
  const results = [];

  for (let i = 0; i < items.length; i++) {
    const item = items[i] || {};
    const sub = item.subscription;
    const endpoint = (sub && sub.endpoint) || item.endpoint || '';

    if (!sub || !sub.endpoint || !sub.keys || !sub.keys.p256dh || !sub.keys.auth) {
      failed++;
      results.push({
        id: item.id || '',
        endpoint: endpoint,
        ok: false,
        expired: false,
        statusCode: 400,
        error: 'Subscription incompleta (endpoint, p256dh e auth são obrigatórios).'
      });
      continue;
    }

    const payloadJson = JSON.stringify(item.payload || {
      title: 'SysManager',
      body: 'Você tem atualizações na sua esteira.'
    });
    const ttlSeconds = Number(item.ttl) > 0 ? Number(item.ttl) : 86400;

    try {
      const resp = await webpush.sendNotification(sub, payloadJson, {
        TTL: ttlSeconds,
        urgency: 'high'
      });
      sent++;
      results.push({
        id: item.id || '',
        endpoint: endpoint,
        ok: true,
        expired: false,
        statusCode: (resp && resp.statusCode) || 201
      });
    } catch (err) {
      failed++;
      const statusCode = Number(err && err.statusCode) || 500;
      const expired = (statusCode === 404 || statusCode === 410);
      if (expired && endpoint && !expiredEndpoints.includes(endpoint)) {
        expiredEndpoints.push(endpoint);
      }
      results.push({
        id: item.id || '',
        endpoint: endpoint,
        ok: false,
        expired: expired,
        statusCode: statusCode,
        error: (err && (err.body || err.message)) || 'Falha ao enviar Web Push'
      });
    }
  }

  return res.status(200).json({
    success: true,
    sent: sent,
    failed: failed,
    expiredEndpoints: expiredEndpoints,
    results: results
  });
};

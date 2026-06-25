const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ROOT_DIR = __dirname;
const DATA_DIR = path.join(ROOT_DIR, 'data');
const LOCAL_TEMPLATES_FILE = path.join(DATA_DIR, 'templates_catalog.json');

const SUPABASE_URL = process.env.SUPABASE_URL || '';
const SUPABASE_ANON_KEY = process.env.SUPABASE_ANON_KEY || '';
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || '';
const SUPABASE_TEMPLATES_TABLE = process.env.SUPABASE_TEMPLATES_TABLE || 'shared_templates';

app.use(express.json({ limit: '2mb' }));

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  if (req.method === 'OPTIONS') return res.status(204).end();
  next();
});

const hasSupabaseConfig = () => Boolean(SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY);

const parseBearerToken = (req) => {
  const auth = req?.headers?.authorization || '';
  const match = String(auth).match(/^Bearer\s+(.+)$/i);
  return match ? match[1].trim() : '';
};

const resolveUser = async (req) => {
  const token = parseBearerToken(req);
  if (!token || !SUPABASE_URL) return null;
  const response = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    method: 'GET',
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${token}`
    }
  });
  if (!response.ok) return null;
  const user = await response.json().catch(() => null);
  return user?.id ? { id: String(user.id) } : null;
};

const ensureLocalTemplatesFile = () => {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  if (!fs.existsSync(LOCAL_TEMPLATES_FILE)) {
    fs.writeFileSync(
      LOCAL_TEMPLATES_FILE,
      JSON.stringify({ schemaVersion: 1, updatedAt: new Date().toISOString(), templates: [] }, null, 2),
      'utf8'
    );
  }
};

const normalizeTemplatesPayload = (templates) => ({
  schemaVersion: 1,
  updatedAt: new Date().toISOString(),
  templates: Array.isArray(templates) ? templates.map((t) => ({
    id: String(t?.id || ''),
    name: String(t?.name || ''),
    data: t?.data && typeof t.data === 'object' ? t.data : {},
    visibility: t?.visibility === 'private' ? 'private' : 'shared',
    ownerId: t?.ownerId ? String(t.ownerId) : (t?.owner_id ? String(t.owner_id) : null)
  })) : []
});

const readLocalTemplates = () => {
  ensureLocalTemplatesFile();
  const raw = fs.readFileSync(LOCAL_TEMPLATES_FILE, 'utf8');
  const parsed = JSON.parse(raw);
  return normalizeTemplatesPayload(parsed?.templates || []);
};

const writeLocalTemplates = (templates) => {
  const payload = normalizeTemplatesPayload(templates);
  ensureLocalTemplatesFile();
  fs.writeFileSync(LOCAL_TEMPLATES_FILE, JSON.stringify(payload, null, 2), 'utf8');
  return payload;
};

const supabaseHeaders = () => ({
  apikey: SUPABASE_SERVICE_ROLE_KEY,
  Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`
});

const readSupabaseTemplates = async (user) => {
  const filter = user?.id
    ? `or=(visibility.eq.shared,and(visibility.eq.private,owner_id.eq.${encodeURIComponent(user.id)}))`
    : `visibility.eq.shared`;
  const url = `${SUPABASE_URL}/rest/v1/${SUPABASE_TEMPLATES_TABLE}?select=id,name,data,updated_at,visibility,owner_id&${filter}&order=updated_at.desc`;
  const response = await fetch(url, { headers: supabaseHeaders() });
  if (!response.ok) throw new Error(`supabase_read_failed_${response.status}`);
  const rows = await response.json();
  const templates = (Array.isArray(rows) ? rows : []).map((row) => ({
    id: String(row?.id || ''),
    name: String(row?.name || ''),
    data: row?.data && typeof row.data === 'object' ? row.data : {},
    visibility: row?.visibility === 'private' ? 'private' : 'shared',
    ownerId: row?.owner_id ? String(row.owner_id) : null
  }));
  return normalizeTemplatesPayload(templates);
};

const fetchSupabaseTemplateIdsForScope = async ({ visibility, user }) => {
  const params = [`select=id`, `visibility=eq.${visibility}`];
  if (visibility === 'private') {
    if (!user?.id) return [];
    params.push(`owner_id=eq.${encodeURIComponent(user.id)}`);
  }
  const url = `${SUPABASE_URL}/rest/v1/${SUPABASE_TEMPLATES_TABLE}?${params.join('&')}`;
  const response = await fetch(url, { headers: supabaseHeaders() });
  if (!response.ok) throw new Error(`supabase_delete_read_failed_${response.status}`);
  const rows = await response.json();
  return (Array.isArray(rows) ? rows : []).map((row) => String(row?.id || '')).filter(Boolean);
};

const deleteSupabaseTemplateIds = async (ids) => {
  const uniqueIds = Array.from(new Set((Array.isArray(ids) ? ids : []).map((id) => String(id || '')).filter(Boolean)));
  if (uniqueIds.length === 0) return;
  const quoted = uniqueIds.map((id) => encodeURIComponent(`"${String(id).replace(/"/g, '\\"')}"`)).join(',');
  const url = `${SUPABASE_URL}/rest/v1/${SUPABASE_TEMPLATES_TABLE}?id=in.(${quoted})`;
  const response = await fetch(url, {
    method: 'DELETE',
    headers: supabaseHeaders()
  });
  if (!response.ok) throw new Error(`supabase_delete_failed_${response.status}`);
};

const deleteMissingSupabaseTemplatesForScope = async ({ templates, visibility, user }) => {
  const incomingIds = new Set((Array.isArray(templates) ? templates : [])
    .filter((t) => (t?.visibility === 'private' ? 'private' : 'shared') === visibility)
    .map((t) => String(t?.id || ''))
    .filter(Boolean));
  const existingIds = await fetchSupabaseTemplateIdsForScope({ visibility, user });
  await deleteSupabaseTemplateIds(existingIds.filter((id) => !incomingIds.has(id)));
};

const writeSupabaseTemplates = async (templates, user) => {
  const payload = normalizeTemplatesPayload(templates).templates;
  await deleteMissingSupabaseTemplatesForScope({ templates: payload, visibility: 'shared', user });
  await deleteMissingSupabaseTemplatesForScope({ templates: payload, visibility: 'private', user });

  const url = `${SUPABASE_URL}/rest/v1/${SUPABASE_TEMPLATES_TABLE}?on_conflict=id`;
  const rows = payload.map((template) => ({
    id: template.id,
    name: template.name,
    data: template.data,
    visibility: template.visibility === 'private' ? 'private' : 'shared',
    owner_id: template.visibility === 'private' ? (user?.id || null) : null
  }));
  if (rows.length === 0) return;
  const response = await fetch(url, {
    method: 'POST',
    headers: {
      ...supabaseHeaders(),
      'Content-Type': 'application/json',
      Prefer: 'resolution=merge-duplicates'
    },
    body: JSON.stringify(rows)
  });
  if (!response.ok) throw new Error(`supabase_write_failed_${response.status}`);
};

app.get('/api/health', (_req, res) => {
  res.json({ ok: true, service: 'calk-v4-railway', ts: new Date().toISOString() });
});

app.get('/api/runtime-config', (_req, res) => {
  res.json({
    ok: true,
    data: {
      SUPABASE_URL: SUPABASE_URL || '',
      SUPABASE_ANON_KEY: SUPABASE_ANON_KEY || ''
    }
  });
});

app.get('/api/templates', async (req, res) => {
  try {
    const user = hasSupabaseConfig() ? await resolveUser(req) : null;
    const data = hasSupabaseConfig() ? await readSupabaseTemplates(user) : readLocalTemplates();
    return res.json({ ok: true, data });
  } catch (error) {
    return res.status(500).json({ ok: false, error: error?.message || 'read_failed' });
  }
});

app.post('/api/templates', async (req, res) => {
  try {
    const templates = Array.isArray(req?.body?.templates) ? req.body.templates : null;
    if (!templates) return res.status(400).json({ ok: false, error: 'invalid_templates_payload' });

    if (hasSupabaseConfig()) {
      const user = await resolveUser(req);
      await writeSupabaseTemplates(templates, user);
    } else {
      writeLocalTemplates(templates);
    }
    return res.json({ ok: true });
  } catch (error) {
    return res.status(500).json({ ok: false, error: error?.message || 'write_failed' });
  }
});

app.use(express.static(ROOT_DIR));

app.get('*', (req, res) => {
  if (req.path.startsWith('/api/')) {
    return res.status(404).json({ ok: false, error: 'not_found' });
  }
  try {
    const indexPath = path.join(ROOT_DIR, 'index.html');
    let html = fs.readFileSync(indexPath, 'utf8');
    const cfgJson = JSON.stringify({
      SUPABASE_URL: SUPABASE_URL || '',
      SUPABASE_ANON_KEY: SUPABASE_ANON_KEY || ''
    });
    html = html.replace(
      /window\.__APP_CONFIG__\s*=\s*window\.__APP_CONFIG__\s*\|\|\s*\{[\s\S]*?\};/,
      `window.__APP_CONFIG__ = Object.assign({}, window.__APP_CONFIG__ || {}, ${cfgJson});`
    );
    return res.type('html').send(html);
  } catch (error) {
    return res.status(500).send('index_render_failed');
  }
});

app.listen(PORT, () => {
  console.log(`Calk v4 server is running on port ${PORT}`);
});

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const TABLE_NAME = process.env.SUPABASE_TEMPLATES_TABLE || 'shared_templates';

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'Content-Type, Authorization',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Content-Type': 'application/json; charset=utf-8'
};

const json = (statusCode, body) => ({
  statusCode,
  headers: corsHeaders,
  body: JSON.stringify(body)
});

const validateConfig = () => {
  if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
    return 'Missing SUPABASE_URL or SUPABASE_SERVICE_ROLE_KEY';
  }
  return null;
};

const parseBearerToken = (event) => {
  const auth = event?.headers?.authorization || event?.headers?.Authorization || '';
  const match = String(auth).match(/^Bearer\s+(.+)$/i);
  return match ? match[1].trim() : '';
};

const resolveUser = async (event) => {
  const token = parseBearerToken(event);
  if (!token) return null;
  const res = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    method: 'GET',
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${token}`
    }
  });
  if (!res.ok) return null;
  const user = await res.json().catch(() => null);
  if (!user?.id) return null;
  return { id: String(user.id) };
};

const fetchTemplates = async (user) => {
  const filter = user?.id
    ? `or=(visibility.eq.shared,and(visibility.eq.private,owner_id.eq.${encodeURIComponent(user.id)}))`
    : `visibility.eq.shared`;
  const url = `${SUPABASE_URL}/rest/v1/${TABLE_NAME}?select=id,name,data,updated_at,visibility,owner_id&${filter}&order=updated_at.desc`;
  const res = await fetch(url, {
    method: 'GET',
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`
    }
  });
  if (!res.ok) throw new Error(`supabase_read_failed_${res.status}`);
  const rows = await res.json();
  return (Array.isArray(rows) ? rows : []).map((row) => ({
    id: String(row.id || ''),
    name: String(row.name || ''),
    data: row.data && typeof row.data === 'object' ? row.data : {},
    visibility: row.visibility === 'private' ? 'private' : 'shared',
    ownerId: row.owner_id ? String(row.owner_id) : null
  }));
};

const upsertTemplates = async (templates, user) => {
  const url = `${SUPABASE_URL}/rest/v1/${TABLE_NAME}?on_conflict=id`;
  const payload = templates.map((t) => ({
    id: String(t.id || ''),
    name: String(t.name || ''),
    data: t.data && typeof t.data === 'object' ? t.data : {},
    visibility: t.visibility === 'private' ? 'private' : 'shared',
    owner_id: t.visibility === 'private' ? (user?.id || null) : null
  }));
  const res = await fetch(url, {
    method: 'POST',
    headers: {
      apikey: SUPABASE_SERVICE_ROLE_KEY,
      Authorization: `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
      'Content-Type': 'application/json',
      Prefer: 'resolution=merge-duplicates'
    },
    body: JSON.stringify(payload)
  });
  if (!res.ok) throw new Error(`supabase_write_failed_${res.status}`);
};

exports.handler = async (event) => {
  if (event.httpMethod === 'OPTIONS') return { statusCode: 204, headers: corsHeaders, body: '' };

  const configError = validateConfig();
  if (configError) return json(500, { ok: false, error: configError });

  try {
    const user = await resolveUser(event);

    if (event.httpMethod === 'GET') {
      const templates = await fetchTemplates(user);
      return json(200, {
        ok: true,
        data: {
          schemaVersion: 1,
          updatedAt: new Date().toISOString(),
          templates
        }
      });
    }

    if (event.httpMethod === 'POST') {
      const parsed = JSON.parse(event.body || '{}');
      const templates = Array.isArray(parsed.templates) ? parsed.templates : null;
      if (!templates) return json(400, { ok: false, error: 'invalid_templates_payload' });
      await upsertTemplates(templates, user);
      return json(200, { ok: true });
    }

    return json(405, { ok: false, error: 'method_not_allowed' });
  } catch (error) {
    return json(500, { ok: false, error: error?.message || 'server_error' });
  }
};

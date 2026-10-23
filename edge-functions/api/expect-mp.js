// edge-functions/api/expect-mp.js
// EdgeOne Pages Edge Function (V8 runtime). KV is bound as the global variable `ttm_kv`.
// GET  -> {Model: mp} map. Source: KV keys "mp:<Model>".
//         If KV empty/unbound, fall back to the static seed /TTM_expect_mp.json.
// POST -> verify "Authorization: Bearer <EXPECT_MP_TOKEN>" (context.env),
//         then write the submitted partial {Model: mp} map (per-Model key, no cross-NPM race).

const KV_PREFIX = 'mp:';
const SEED_URL = '/TTM_expect_mp.json';

function json(data, status) {
  return new Response(JSON.stringify(data), {
    status: status || 200,
    headers: {
      'Content-Type': 'application/json; charset=UTF-8',
      'Access-Control-Allow-Origin': '*',
      'Cache-Control': 'no-store',
    },
  });
}

async function kvMap() {
  const out = {};
  let cursor = undefined;
  // ttm_kv is injected by the KV binding as a global; throws if not bound.
  do {
    const page = await ttm_kv.list({ prefix: KV_PREFIX, limit: 1000, cursor: cursor });
    for (const k of page.keys) {
      const model = k.name.slice(KV_PREFIX.length);
      const v = await ttm_kv.get(k.name);
      if (v != null) out[model] = v;
    }
    cursor = page.cursor;
  } while (cursor);
  return out;
}

export async function onRequestGet(context) {
  try {
    const map = await kvMap();
    if (Object.keys(map).length === 0) {
      const r = await fetch(SEED_URL, { cache: 'no-store' });
      if (r.ok) return json(await r.json());
    }
    return json(map);
  } catch (e) {
    // KV unbound or error -> try static seed.
    try {
      const r = await fetch(SEED_URL, { cache: 'no-store' });
      if (r.ok) return json(await r.json());
    } catch (_) {}
    return json({ error: String(e) }, 500);
  }
}

export async function onRequestPost(context) {
  try {
    const env = context.env || {};
    const expected = env.EXPECT_MP_TOKEN;
    const auth = context.request.headers.get('Authorization') || '';
    if (!expected || auth !== 'Bearer ' + expected) {
      return json({ error: 'unauthorized' }, 401);
    }
    const body = await context.request.json();
    if (!body || typeof body !== 'object') {
      return json({ error: 'bad body' }, 400);
    }
    let count = 0;
    for (const model of Object.keys(body)) {
      const val = body[model];
      if (val === '' || val == null) {
        await ttm_kv.delete(KV_PREFIX + model);
      } else {
        await ttm_kv.put(KV_PREFIX + model, String(val));
        count++;
      }
    }
    return json({ ok: true, count: count });
  } catch (e) {
    return json({ error: String(e) }, 500);
  }
}

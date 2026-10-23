// functions/api/expect-mp.js
// EdgeOne Pages Function (V8 runtime). Route: /api/expect-mp
// KV namespace is bound to the project; the binding variable name is `ttm_kv`
// (accessed as a global, with context.env.ttm_kv as fallback).
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

// KV binding: prefer the injected global `ttm_kv`; fall back to context.env
// (typeof guard is safe for undeclared identifiers).
function getKV(context) {
  try {
    if (typeof ttm_kv !== 'undefined' && ttm_kv) return ttm_kv;
  } catch (e) {}
  const env = (context && context.env) || {};
  if (env.ttm_kv) return env.ttm_kv;
  return null;
}

async function kvMap(context) {
  const kv = getKV(context);
  if (!kv) throw new Error('KV not bound');
  const out = {};
  let cursor = undefined;
  do {
    const page = await kv.list({ prefix: KV_PREFIX, limit: 1000, cursor: cursor });
    for (const k of page.keys) {
      const model = k.name.slice(KV_PREFIX.length);
      const v = await kv.get(k.name);
      if (v != null) out[model] = v;
    }
    cursor = page.cursor;
  } while (cursor);
  return out;
}

export async function onRequestGet(context) {
  try {
    const map = await kvMap(context);
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
    const kv = getKV(context);
    if (!kv) return json({ error: 'KV not bound' }, 500);
    let count = 0;
    for (const model of Object.keys(body)) {
      const val = body[model];
      if (val === '' || val == null) {
        await kv.delete(KV_PREFIX + model);
      } else {
        await kv.put(KV_PREFIX + model, String(val));
        count++;
      }
    }
    return json({ ok: true, count: count });
  } catch (e) {
    return json({ error: String(e) }, 500);
  }
}

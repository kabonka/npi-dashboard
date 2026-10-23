// functions/api/expect-mp.js
// EdgeOne Pages Function (V8 runtime). Route: /api/expect-mp
// KV namespace is bound to the project; the binding variable name is `ttm_kv`
// (accessed as a global, with context.env.ttm_kv as fallback).
// GET  -> {Model: mp} map. Source: KV keys "mp_<hex(Model)>".
//         If KV empty/unbound, fall back to the static seed /TTM_expect_mp.json.
// POST -> verify "Authorization: Bearer <EXPECT_MP_TOKEN>" (context.env),
//         then write the submitted partial {Model: mp} map (per-Model key, no cross-NPM race).
//
// KV constraints (official doc):
//   - key: <=512B, ONLY [0-9a-zA-Z_] -> Model names are hex-encoded.
//   - list({prefix, limit<=256, cursor}) -> {complete, cursor, keys:[{key}]}

const KV_PREFIX = 'mp_';
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

// Model -> fixed 4-hex per UTF-16 unit so the KV key only contains [0-9a-zA-Z_].
function encModel(m) {
  let s = '';
  for (const ch of String(m)) {
    s += ch.charCodeAt(0).toString(16).padStart(4, '0');
  }
  return s;
}

function decModel(h) {
  let s = '';
  for (let i = 0; i + 3 < h.length; i += 4) {
    s += String.fromCharCode(parseInt(h.slice(i, i + 4), 16));
  }
  return s;
}

async function kvMap(context) {
  const kv = getKV(context);
  if (!kv) throw new Error('KV not bound');
  const out = {};
  let cursor;
  for (;;) {
    const page = await kv.list({ prefix: KV_PREFIX, limit: 256, cursor: cursor });
    for (const k of (page.keys || [])) {
      const key = k.key || k.name; // doc says .key; .name as defensive fallback
      if (!key) continue;
      const model = decModel(key.slice(KV_PREFIX.length));
      const v = await kv.get(key);
      if (v != null) out[model] = v;
    }
    if (page.complete || !page.cursor) break;
    cursor = page.cursor;
  }
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
      const key = KV_PREFIX + encModel(model);
      if (val === '' || val == null) {
        await kv.delete(key);
      } else {
        await kv.put(key, String(val));
        count++;
      }
    }
    return json({ ok: true, count: count });
  } catch (e) {
    return json({ error: String(e) }, 500);
  }
}

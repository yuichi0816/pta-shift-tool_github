export async function onRequestPost({ env, request }) {
  const kv = env.REGION_KV;

  // 合言葉チェック（?token=...）
  const url = new URL(request.url);
  const token = url.searchParams.get("token") || "";
  if (!env.RESET_TOKEN || token !== env.RESET_TOKEN) {
    return new Response("Unauthorized", { status: 401 });
  }

  // KVは list が分割されるので cursor で全件削除
  const prefixes = ["day:", "total:"];

  for (const prefix of prefixes) {
    let cursor = undefined;
    while (true) {
      const res = await kv.list({ prefix, cursor });
      for (const k of res.keys) {
        await kv.delete(k.name);
      }
      if (!res.list_complete) {
        cursor = res.cursor;
      } else {
        break;
      }
    }
  }

  return new Response("OK", { status: 200 });
}

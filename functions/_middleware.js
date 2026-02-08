function ymdTokyo() {
  const fmt = new Intl.DateTimeFormat("ja-JP", {
    timeZone: "Asia/Tokyo",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });
  const p = Object.fromEntries(fmt.formatToParts(new Date()).map(x => [x.type, x.value]));
  return `${p.year}-${p.month}-${p.day}`;
}

export async function onRequest(context) {
  const kv = context.env && context.env.REGION_KV;
  if (!kv) return context.next();

  const day = ymdTokyo();
  const cookieName = `counted_${day}`;

  // ① その日のCookieがあればカウントしない
  const cookie = context.request.headers.get("Cookie") || "";
  if (cookie.includes(`${cookieName}=1`)) {
    return context.next();
  }

  // ② まだならカウント
  const cf = context.request.cf || {};
  const country = cf.country || "XX";   // ★追加（JP/US等）
  const region = cf.region || "Unknown";



  // 全体（日別）
  const totalKey = `total:${day}`;
  const total = Number(await kv.get(totalKey)) || 0;
  await kv.put(totalKey, String(total + 1));

  // 地域（日別）
  // 例：day:2025-12-30:country:JP:region:Chiba
  const key = `day:${day}:country:${country}:region:${region}`;
  const cur = Number(await kv.get(key)) || 0;
  await kv.put(key, String(cur + 1));

  // ③ レスポンスにCookieを付与（その日限り）
  const res = await context.next();
  const expire = new Date(`${day}T23:59:59+09:00`).toUTCString();

  res.headers.append(
    "Set-Cookie",
    `${cookieName}=1; Path=/; Expires=${expire}; SameSite=Lax`
  );

  return res;
}

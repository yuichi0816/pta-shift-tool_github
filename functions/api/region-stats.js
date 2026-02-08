// 英語地域名 → 日本語都道府県名のマッピング
const PREFECTURE_MAP = {
  // 北海道・東北
  "Hokkaido": "北海道",
  "Aomori": "青森県",
  "Iwate": "岩手県",
  "Miyagi": "宮城県",
  "Akita": "秋田県",
  "Yamagata": "山形県",
  "Fukushima": "福島県",
  // 関東
  "Ibaraki": "茨城県",
  "Tochigi": "栃木県",
  "Gunma": "群馬県",
  "Saitama": "埼玉県",
  "Chiba": "千葉県",
  "Tokyo": "東京都",
  "Kanagawa": "神奈川県",
  // 中部
  "Niigata": "新潟県",
  "Toyama": "富山県",
  "Ishikawa": "石川県",
  "Fukui": "福井県",
  "Yamanashi": "山梨県",
  "Nagano": "長野県",
  "Gifu": "岐阜県",
  "Shizuoka": "静岡県",
  "Aichi": "愛知県",
  // 近畿
  "Mie": "三重県",
  "Shiga": "滋賀県",
  "Kyoto": "京都府",
  "Osaka": "大阪府",
  "Hyogo": "兵庫県",
  "Nara": "奈良県",
  "Wakayama": "和歌山県",
  // 中国
  "Tottori": "鳥取県",
  "Shimane": "島根県",
  "Okayama": "岡山県",
  "Hiroshima": "広島県",
  "Yamaguchi": "山口県",
  // 四国
  "Tokushima": "徳島県",
  "Kagawa": "香川県",
  "Ehime": "愛媛県",
  "Kochi": "高知県",
  // 九州・沖縄
  "Fukuoka": "福岡県",
  "Saga": "佐賀県",
  "Nagasaki": "長崎県",
  "Kumamoto": "熊本県",
  "Oita": "大分県",
  "Miyazaki": "宮崎県",
  "Kagoshima": "鹿児島県",
  "Okinawa": "沖縄県",
  // 不明
  "Unknown": "不明"
};

// 国コード → 国名のマッピング（主要な国のみ）
const COUNTRY_MAP = {
  "JP": "日本",
  "US": "アメリカ",
  "CN": "中国",
  "KR": "韓国",
  "TW": "台湾",
  "HK": "香港",
  "SG": "シンガポール",
  "TH": "タイ",
  "VN": "ベトナム",
  "PH": "フィリピン",
  "ID": "インドネシア",
  "MY": "マレーシア",
  "IN": "インド",
  "AU": "オーストラリア",
  "NZ": "ニュージーランド",
  "GB": "イギリス",
  "DE": "ドイツ",
  "FR": "フランス",
  "IT": "イタリア",
  "ES": "スペイン",
  "NL": "オランダ",
  "CA": "カナダ",
  "BR": "ブラジル",
  "RU": "ロシア",
  "XX": "不明"
};

// 英語地域名を日本語都道府県名に変換
function toJapanesePrefecture(englishRegion) {
  return PREFECTURE_MAP[englishRegion] || englishRegion;
}

// 国コードを国名に変換
function toCountryName(countryCode) {
  return COUNTRY_MAP[countryCode] || countryCode;
}

export async function onRequestGet({ env, request }) {
  const kv = env.REGION_KV;

  const url = new URL(request.url);
  const day = url.searchParams.get("day");
  if (!day) {
    return new Response(
      JSON.stringify({ error: "day is required. e.g. /api/region-stats?day=2025-12-29" }),
      { status: 400, headers: { "content-type": "application/json; charset=utf-8" } }
    );
  }

  // 新形式: day:YYYY-MM-DD:country:JP:region:Chiba
  const newPrefix = `day:${day}:country:`;
  const newList = await kv.list({ prefix: newPrefix });

  // 旧形式も検索: day:YYYY-MM-DD:region:JP-13:Tokyo
  const oldPrefix = `day:${day}:region:`;
  const oldList = await kv.list({ prefix: oldPrefix });

  const rows = [];

  // 新形式のキーを処理
  for (const k of newList.keys) {
    const count = Number(await kv.get(k.name)) || 0;
    const parts = k.name.split(":");
    // key: day:YYYY-MM-DD:country:JP:region:Chiba
    const country = parts[3] || "XX";
    const region = parts[5] || "Unknown";

    // 日本国内・海外の判定と表示名の設定
    const isJapan = (country === "JP");
    const displayRegion = isJapan
      ? toJapanesePrefecture(region)  // 日本: 都道府県名
      : toCountryName(country);        // 海外: 国名
    const displayType = isJapan ? "国内" : "海外";

    rows.push({
      country,
      region,
      displayRegion,
      displayType,
      countryName: toCountryName(country),
      prefectureName: isJapan ? toJapanesePrefecture(region) : null,
      count
    });
  }

  // 旧形式のキーを処理
  for (const k of oldList.keys) {
    const count = Number(await kv.get(k.name)) || 0;
    const parts = k.name.split(":");
    // key: day:YYYY-MM-DD:region:JP-13:Tokyo
    const regionCode = parts[3] || "";
    const regionName = parts[4] || "Unknown";
    // JP-XX の形式なら日本と判定
    const isJapan = regionCode.startsWith("JP-");
    const country = isJapan ? "JP" : "XX";

    const displayRegion = isJapan
      ? toJapanesePrefecture(regionName)
      : toCountryName(country);
    const displayType = isJapan ? "国内" : "海外";

    rows.push({
      country,
      region: regionName,
      displayRegion,
      displayType,
      countryName: toCountryName(country),
      prefectureName: isJapan ? toJapanesePrefecture(regionName) : null,
      count
    });
  }

  rows.sort((a, b) => b.count - a.count);

  const rawTotal = await kv.get(`total:${day}`);
  const total = rawTotal ? Number(rawTotal) : 0;

  return new Response(JSON.stringify({ day, total, rows }), {
    headers: { "content-type": "application/json; charset=utf-8" },
  });
}

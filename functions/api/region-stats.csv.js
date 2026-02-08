// 英語地域名 → 日本語都道府県名のマッピング
const PREFECTURE_MAP = {
  "Hokkaido": "北海道",
  "Aomori": "青森県",
  "Iwate": "岩手県",
  "Miyagi": "宮城県",
  "Akita": "秋田県",
  "Yamagata": "山形県",
  "Fukushima": "福島県",
  "Ibaraki": "茨城県",
  "Tochigi": "栃木県",
  "Gunma": "群馬県",
  "Saitama": "埼玉県",
  "Chiba": "千葉県",
  "Tokyo": "東京都",
  "Kanagawa": "神奈川県",
  "Niigata": "新潟県",
  "Toyama": "富山県",
  "Ishikawa": "石川県",
  "Fukui": "福井県",
  "Yamanashi": "山梨県",
  "Nagano": "長野県",
  "Gifu": "岐阜県",
  "Shizuoka": "静岡県",
  "Aichi": "愛知県",
  "Mie": "三重県",
  "Shiga": "滋賀県",
  "Kyoto": "京都府",
  "Osaka": "大阪府",
  "Hyogo": "兵庫県",
  "Nara": "奈良県",
  "Wakayama": "和歌山県",
  "Tottori": "鳥取県",
  "Shimane": "島根県",
  "Okayama": "岡山県",
  "Hiroshima": "広島県",
  "Yamaguchi": "山口県",
  "Tokushima": "徳島県",
  "Kagawa": "香川県",
  "Ehime": "愛媛県",
  "Kochi": "高知県",
  "Fukuoka": "福岡県",
  "Saga": "佐賀県",
  "Nagasaki": "長崎県",
  "Kumamoto": "熊本県",
  "Oita": "大分県",
  "Miyazaki": "宮崎県",
  "Kagoshima": "鹿児島県",
  "Okinawa": "沖縄県",
  "Unknown": "不明"
};

// 国コード → 国名のマッピング
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

function toJapanesePrefecture(englishRegion) {
  return PREFECTURE_MAP[englishRegion] || englishRegion;
}

function toCountryName(countryCode) {
  return COUNTRY_MAP[countryCode] || countryCode;
}

export async function onRequestGet({ env, request }) {
  const kv = env.REGION_KV;

  const url = new URL(request.url);
  const day = url.searchParams.get("day");
  if (!day) {
    return new Response(
      "day is required. e.g. /api/region-stats.csv?day=2025-12-29",
      { status: 400 }
    );
  }

  const prefix = `day:${day}:country:`;
  const list = await kv.list({ prefix });

  const rows = [];
  for (const k of list.keys) {
    const count = Number(await kv.get(k.name)) || 0;
    const parts = k.name.split(":");
    // key: day:YYYY-MM-DD:country:JP:region:Chiba
    const country = parts[3] || "XX";
    const region = parts[5] || "Unknown";

    // 国内/海外の判定
    const isJapan = (country === "JP");
    const displayType = isJapan ? "国内" : "海外";
    const displayRegion = isJapan
      ? toJapanesePrefecture(region)  // 都道府県名
      : toCountryName(country);        // 国名

    rows.push([day, displayType, country, toCountryName(country), region, displayRegion, count]);
  }

  rows.sort((a, b) => b[6] - a[6]);

  const csv = [
    ["day", "区分", "country_code", "国名", "region_code", "地域名", "count"],
    ...rows,
  ]
    .map(row => row.map(v => `"${String(v).replaceAll(`"`, `""`)}"`).join(","))
    .join("\n");

  return new Response(csv, {
    headers: {
      "content-type": "text/csv; charset=utf-8",
      "content-disposition": `attachment; filename="region-stats-${day}.csv"`,
    },
  });
}


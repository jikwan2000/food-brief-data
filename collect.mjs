// 식자재 브리프 데이터 수집기 — GitHub Actions에서 매일 실행되어 data.json을 갱신합니다.
import fs from 'node:fs/promises';
import * as XLSX from 'xlsx';

const OUT = 'data.json';
const prev = await fs.readFile(OUT, 'utf8').then(JSON.parse).catch(() => ({ items: {}, news: [] }));
const out = { updated: new Date().toISOString(), items: { ...prev.items }, news: prev.news || [], errors: [] };
const UA = { 'User-Agent': 'Mozilla/5.0 (food-brief-collector)' };
const get = async (u) => { const r = await fetch(u, { headers: UA }); if (!r.ok) throw new Error(`${u} → ${r.status}`); return r; };
const iso = (d) => d.toISOString().slice(0, 10);
const mean = (a) => a.reduce((s, x) => s + x, 0) / a.length;
// 평년: 최근 5년 값 중 최고·최저를 뺀 3년 평균 (5년이 안 되면 단순 평균)
const olympic = (a) => !a.length ? null : a.length >= 5 ? mean([...a].sort((x, y) => x - y).slice(1, -1)) : mean(a);
async function step(name, fn) {
  try { await fn(); console.log('ok  ', name); }
  catch (e) { console.error('FAIL', name, e.message); out.errors.push(`${name}: ${e.message}`); }
}

// 1) 뉴스 — 구글 뉴스 RSS (최근 7일)
const QUERIES = [
  { q: '배추 가격', region: '국내', cat: '농산물', item: 'cabbage' },
  { q: '농산물 도매가격', region: '국내', cat: '농산물' },
  { q: '고등어 가격', region: '국내', cat: '수산물', item: 'mackerel' },
  { q: '수산물 가격', region: '국내', cat: '수산물' },
  { q: '계란 가격 OR 돼지고기 가격 OR 한우 가격', region: '국내', cat: '축산물' },
  { q: '농축산물 물가 대책', region: '국내', cat: '정책·유통' },
  { q: '국제 설탕 가격 OR 원당', region: '해외', cat: '가공·소재', item: 'rawsugar' },
  { q: '국제 밀 가격 OR 소맥', region: '해외', cat: '가공·소재', item: 'wheat' },
  { q: '국제 옥수수 가격', region: '해외', cat: '가공·소재', item: 'corn' },
  { q: '팜유 OR 대두유 가격', region: '해외', cat: '가공·소재', item: 'palm' },
  { q: '국제 쌀 가격', region: '해외', cat: '가공·소재', item: 'thairice' },
  { q: 'GDT 분유 OR 유제품 국제가격', region: '해외', cat: '가공·소재', item: 'wmp' },
  { q: 'FAO food price index', region: '해외', cat: '가공·소재', en: true, n: 2 },
  { q: 'sugar prices', region: '해외', cat: '가공·소재', item: 'rawsugar', en: true, n: 2 },
  { q: 'wheat prices', region: '해외', cat: '가공·소재', item: 'wheat', en: true, n: 2 },
  { q: 'Global Dairy Trade auction', region: '해외', cat: '가공·소재', item: 'wmp', en: true, n: 2 },
];
await step('news', async () => {
  const seen = new Set(), list = [];
  const clean = (s) => s.replace(/<!\[CDATA\[|\]\]>/g, '').replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').trim();
  for (const q of QUERIES) {
    const lang = q.en ? 'hl=en-US&gl=US&ceid=US:en' : 'hl=ko&gl=KR&ceid=KR:ko';
    const xml = await (await get(`https://news.google.com/rss/search?q=${encodeURIComponent(q.q + ' when:7d')}&${lang}`)).text();
    for (const [, b] of [...xml.matchAll(/<item>([\s\S]*?)<\/item>/g)].slice(0, q.n || 4)) {
      const tag = (t) => (b.match(new RegExp(`<${t}[^>]*>([\\s\\S]*?)<\\/${t}>`)) || [])[1] || '';
      const source = clean(tag('source'));
      let title = clean(tag('title'));
      if (source && title.endsWith(' - ' + source)) title = title.slice(0, -(source.length + 3));
      const key = title.replace(/\s/g, '').slice(0, 30);
      if (!title || seen.has(key)) continue;
      seen.add(key);
      list.push({ id: 'g' + list.length, d: iso(new Date(clean(tag('pubDate')))), region: q.region, cat: q.cat, item: q.item || null,
        source: source + (q.en ? ' (영문)' : ''), url: clean(tag('link')), title, summary: '' });
    }
  }
  if (!list.length) throw new Error('기사 0건');
  out.news = list.sort((a, b) => b.d.localeCompare(a.d)).slice(0, 40);
});

// 2) 배추 — 가락시장 경매가 (중품, 10kg 망대) + 최근 5년 같은 시기 평년
await step('cabbage', async () => {
  const Y = new Date().getFullYear(), years = {};
  for (let y = Y - 5; y <= Y; y++) {
    const html = await (await get(`https://www.garakprice.com/pum_detail.php?pum_cd=21100&year=${y}&month=&ca1=${encodeURIComponent('중')}`)).text();
    const text = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
    years[y] = [...text.matchAll(/(\d{4})\/(\d{1,2})\/(\d{1,2}) 10 키로망대 중 ([\d,]+)/g)]
      .map((m) => [`${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`, +m[4].replace(/,/g, '')])
      .sort((a, b) => a[0].localeCompare(b[0]));
  }
  const raw = years[Y];
  if (!raw.length) throw new Error('올해 데이터 0건 — 페이지 구조 확인 필요');
  const md = raw[raw.length - 1][0].slice(5);
  const past = [];
  for (let y = Y - 5; y < Y; y++) {
    const t = Date.parse(`${y}-${md}`);
    const v = (years[y] || []).filter(([d]) => Math.abs(Date.parse(d) - t) <= 7 * 864e5).map((x) => x[1]);
    if (v.length) past.push(mean(v));
  }
  out.items.cabbage = { raw, base: olympic(past),
    src: '서울 가락시장 경매가 (garakprice.com) · 자동 수집 · 평년 = 최근 5년 같은 시기(±7일) 평균 중 최고·최저 제외 평균' };
});

// 3) 글로벌 원자재 — World Bank 월간 가격 엑셀 (당·분·유·미, 커피, 코코아)
await step('worldbank', async () => {
  const page = await (await get('https://www.worldbank.org/en/research/commodity-markets')).text();
  const link = page.match(/https?:\/\/[^"'\s]*CMO-Historical-Data-Monthly\.xlsx/i);
  if (!link) throw new Error('월간 엑셀 링크를 찾지 못함');
  const wbk = XLSX.read(Buffer.from(await (await get(link[0])).arrayBuffer()));
  const rows = XLSX.utils.sheet_to_json(wbk.Sheets['Monthly Prices'], { header: 1, raw: true });
  const hi = rows.findIndex((r) => r.some((c) => String(c).trim() === 'Cocoa'));
  const head = Array.from(rows[hi], (c) => String(c ?? '').toLowerCase().trim());
  const data = rows.slice(hi + 1).filter((r) => /^\d{4}M\d{2}$/.test(String(r[0])));
  const MAP = { rawsugar: 'sugar, world', wheat: 'wheat, us hrw', corn: 'maize', soyoil: 'soybean oil', palm: 'palm oil',
    thairice: 'rice, thai 5%', vnrice: 'rice, viet namese 5%', coffee: 'coffee, arabica', cocoa: 'cocoa' };
  for (const [id, name] of Object.entries(MAP)) {
    const ci = head.findIndex((h) => h === name || h.startsWith(name));
    if (ci < 0) { out.errors.push(`worldbank: 열 없음 ${name}`); continue; }
    const ser = data.map((r) => [String(r[0]), Number(r[ci])]).filter(([, v]) => isFinite(v) && v > 0);
    const lastY = +ser[ser.length - 1][0].slice(0, 4);
    const annual = [];
    for (let y = lastY - 5; y < lastY; y++) { const v = ser.filter(([p]) => +p.slice(0, 4) === y).map((x) => x[1]); if (v.length === 12) annual.push(mean(v)); }
    out.items[id] = { freq: 'mm', base: olympic(annual),
      pts: ser.slice(-13).map(([p, v]) => [`${p.slice(2, 4)}.${p.slice(5)}`, `${p.slice(0, 4)}-${p.slice(5)}-15`, v]),
      src: 'World Bank CMO 월간 가격 · 자동 수집 · 평년 = 최근 5년 연평균 중 최고·최저 제외 평균' };
  }
});
// 4) 당·분·유 일별 시세 — 선물 근월물 종가 (Yahoo Finance)
const FUT = {
  rawsugar: ['SB=F', 0.0220462, 4, 'ICE 원당 선물 No.11', '$/kg'],
  wheat:    ['KE=F', 0.367437, 1, 'KC 소맥 선물 (HRW)', '$/t'],
  soyoil:   ['ZL=F', 22.0462, 1, 'CBOT 대두유 선물', '$/t'],
};
await step('futures', async () => {
  for (const [id, [sym, k, dp, market, unit]] of Object.entries(FUT)) {
    let j;
    for (const h of ['query1', 'query2']) {
      try { j = await (await get(`https://${h}.finance.yahoo.com/v8/finance/chart/${encodeURIComponent(sym)}?range=6y&interval=1d`)).json(); break; }
      catch (e) { if (h === 'query2') throw e; }
    }
    const r = j.chart.result[0], cl = r.indicators.quote[0].close;
    const ser = r.timestamp.map((t, i) => [iso(new Date(t * 1000)), cl[i]]).filter(([, v]) => v > 0).map(([d, v]) => [d, +(v * k).toFixed(dp)]);
    if (ser.length < 100) throw new Error(`${sym} 데이터 부족`);
    const lastY = +ser[ser.length - 1][0].slice(0, 4), annual = [];
    for (let y = lastY - 5; y < lastY; y++) { const v = ser.filter(([d]) => +d.slice(0, 4) === y).map((x) => x[1]); if (v.length > 200) annual.push(mean(v)); }
    const from = iso(new Date(Date.now() - 400 * 864e5));
    out.items[id] = { raw: ser.filter(([d]) => d >= from), base: olympic(annual), market, unit,
      src: `${market} 근월물 종가 (Yahoo Finance) · 일별 자동 수집 · 평년 = 최근 5년 연평균 중 최고·최저 제외 평균` };
  }
});
await fs.writeFile(OUT, JSON.stringify(out, null, 1));
console.log(`done · news ${out.news.length} · items ${Object.keys(out.items).length} · errors ${out.errors.length}`);

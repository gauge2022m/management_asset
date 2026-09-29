/**
 * マイ資産台帳 用の価格取得スクリプト（Google Apps Script）
 * 使い方は README の「株価の自動取得を設定する」を参照。
 * 下の TOKEN を好きな合言葉に変えて、アプリの設定にも同じものを入れる。
 */
const TOKEN = 'ここを好きな合言葉に変える';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (TOKEN.indexOf('ここを') === 0) return out({ error: 'スクリプトの合言葉(TOKEN)が未設定です' });
  if (p.token !== TOKEN) return out({ error: '合言葉が違います' });
  const keys = String(p.keys || '').split(',').map(s => s.trim()).filter(String);
  if (p.test) return out({ ok: true, prices: {} });
  const res = {};
  const gf = [];
  keys.forEach(k => {
    const i = k.indexOf(':');
    const kind = k.slice(0, i), sym = k.slice(i + 1);
    if (kind === 'jp') gf.push([k, 'TYO:' + sym, 'JPY']);
    else if (kind === 'us') gf.push([k, sym, 'USD']);
    else if (kind === 'crypto') gf.push([k, 'CURRENCY:' + sym + 'JPY', 'JPY']);
    else if (kind === 'fund') res[k] = fundNav(sym);
  });
  if (gf.length) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('prices') || ss.insertSheet('prices');
    sh.clear();
    const rows = gf.map(g => [g[0], '=GOOGLEFINANCE("' + g[1] + '","price")', g[2]]);
    sh.getRange(1, 1, rows.length, 3).setValues(rows);
    let vals = [];
    // 数式の計算が終わるまで最大10秒ほど待つ
    for (let t = 0; t < 8; t++) {
      SpreadsheetApp.flush();
      Utilities.sleep(t === 0 ? 1500 : 1200);
      vals = sh.getRange(1, 1, rows.length, 3).getValues();
      if (vals.every(r => typeof r[1] === 'number' || /^#/.test(String(r[1])))) break;
    }
    vals.forEach((r, i) => {
      const price = r[1];
      if (typeof price === 'number' && price > 0) {
        res[r[0]] = { price: price, currency: r[2], asOf: new Date().toISOString(), source: 'GOOGLEFINANCE' };
      } else {
        // GOOGLEFINANCE で取れない時は Google Finance のページから読む
        const alt = financePage(gf[i]);
        res[r[0]] = alt || { error: 'GOOGLEFINANCE: ' + (String(price) || '空') + ' / ページからも取得できません' };
      }
    });
  }
  return out({ ok: true, prices: res });
}

/** 投資信託協会のページから基準価額（1万口あたり）を読む */
function fundNav(isin) {
  if (!/^JP[0-9A-Z]{10}$/.test(isin)) return { error: 'ISINコード(JP…)で登録してください' };
  try {
    const html = UrlFetchApp.fetch('https://toushin-lib.fwg.ne.jp/FdsWeb/FDST030000?isinCd=' + isin,
      { muteHttpExceptions: true, followRedirects: true }).getContentText('UTF-8');
    const text = html.replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ');
    const i = text.indexOf('基準価額');
    if (i < 0) return { error: '基準価額が見つかりません' };
    const seg = text.slice(i, i + 400);
    const m = seg.match(/([0-9]{1,3}(?:,[0-9]{3})+|[0-9]{4,6})\s*円/);
    if (!m) return { error: '基準価額を読み取れません' };
    const price = Number(m[1].replace(/,/g, ''));
    const d = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
    const asOf = d ? new Date(Number(d[1]), Number(d[2]) - 1, Number(d[3]), 15).toISOString() : new Date().toISOString();
    const nm = text.match(/ファンド名称?\s*([^\s].{2,60}?)\s/);
    return { price: price, currency: 'JPY', asOf: asOf, source: '投資信託協会', name: nm ? nm[1] : '' };
  } catch (err) {
    return { error: '投資信託協会に接続できません' };
  }
}

/** Google Finance のページ（例 https://www.google.com/finance/quote/6702:TYO）から現在値を読む */
function financePage(g) {
  const k = g[0], kind = k.slice(0, k.indexOf(':')), sym = k.slice(k.indexOf(':') + 1);
  const urls = kind === 'jp' ? [sym + ':TYO']
    : kind === 'crypto' ? [sym + '-JPY']
    : [sym + ':NASDAQ', sym + ':NYSE', sym + ':NYSEARCA', sym + ':BATS'];
  for (const q of urls) {
    try {
      const html = UrlFetchApp.fetch('https://www.google.com/finance/quote/' + q + '?hl=ja',
        { muteHttpExceptions: true, headers: { 'Accept-Language': 'ja' } }).getContentText('UTF-8');
      const m = html.match(/data-last-price="([0-9.]+)"/);
      if (m && Number(m[1]) > 0) {
        return { price: Number(m[1]), currency: g[2], asOf: new Date().toISOString(), source: 'Google Finance' };
      }
    } catch (err) {}
  }
  return null;
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** エディタで一度実行して、権限を許可するための関数 */
function authorize() {
  SpreadsheetApp.getActiveSpreadsheet();
  UrlFetchApp.fetch('https://toushin-lib.fwg.ne.jp/', { muteHttpExceptions: true });
}

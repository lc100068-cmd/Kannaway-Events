/**
 * キャナウェイジャパン社長ロバート・ケリー氏と「麻」が持つ本当の魅力を学ぶ会 お申し込み受付スクリプト
 *
 * Google スプレッドシートの「拡張機能 → Apps Script」に貼り付けて使います。
 * 設定を変えるときは CONFIG だけを書き換えてください。
 */
const CONFIG = {
  // 定員（同伴者を含む合計人数）
  CAPACITY: 40,
  // 申込締切（日本時間）。config.js の DEADLINE と同じにしてください。
  DEADLINE: '2026-11-18T23:59:59+09:00',
  // 申し込みがあったときに通知を受け取るメールアドレス（空なら通知しない）
  ORGANIZER_EMAIL: '',
  // 確認メールの差出人名
  SENDER_NAME: 'ロバート・ケリー氏と学ぶ会 事務局',
  // 確認メールの末尾に載せるお問い合わせ先（空なら載せない）
  CONTACT: '',
  // 申込一覧を書き込むシート名
  SHEET_NAME: '申込一覧',
};

const EVENT = {
  title: 'キャナウェイジャパン社長ロバート・ケリー氏と「麻」が持つ本当の魅力を学ぶ会',
  timeLabel: '受付 12:30〜12:50／13:00〜17:30',
  dateLabel: '2026年11月23日（月・祝）',
  venue: 'レンタルスペース GOODLUCK',
  address: '〒703-8203 岡山県岡山市中区国府市場331-3',
  mapUrl: 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent('岡山県岡山市中区国府市場331-3'),
};

// 「参加する部」列に書き込む値（部の区別はなくなったので一律）
const ENTRY_LABEL = '参加';

const HEADERS = [
  '受付日時', '受付番号', 'お名前', 'ふりがな', '電話番号', 'メールアドレス',
  '参加する部', '同伴者数', '合計人数', '会員区分', 'ご紹介者', '次回案内希望',
  '質問・メッセージ', 'キャンセル（何か入力すると人数から除外）',
];
const COL = { part: 6, total: 8, cancel: 13 }; // 0始まりの列番号

/** 初回に一度だけ実行：シート作成とメール送信の許可 */
function setup() {
  getSheet_();
  MailApp.getRemainingDailyQuota();
}

/** 残席・受付状況を返す */
function doGet() {
  return json_(status_(getSheet_()));
}

/** 申し込みを受け付ける */
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, code: 'invalid', message: '送信内容を読み取れませんでした。' });
  }
  // スパム対策（人には見えない欄に入力があれば保存せず成功扱い）
  if (data.website) return json_({ ok: true, receiptNo: '-' });

  const checked = validate_(data);
  if (checked.error) return json_({ ok: false, code: 'invalid', message: checked.error });
  const entry = checked.entry;

  if (isClosed_()) {
    return json_({ ok: false, code: 'closed', message: 'お申し込みの受付は終了しました。' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, code: 'busy', message: 'ただいま混み合っています。少し時間をおいて、もう一度お試しください。' });
  }

  let receiptNo;
  try {
    const sheet = getSheet_();
    const status = status_(sheet);
    const left = status.remaining.all;
    if (entry.total > left) {
      return json_({
        ok: false,
        code: left > 0 ? 'full' : 'closed',
        message: left > 0
          ? `申し訳ありません。残りのお席は${left}席です。人数をご確認ください。`
          : '申し訳ありません。定員に達したため、受付を終了しました。',
        remaining: status.remaining,
      });
    }

    const props = PropertiesService.getScriptProperties();
    const seq = Number(props.getProperty('seq') || 0) + 1;
    props.setProperty('seq', String(seq));
    receiptNo = 'K' + ('000' + seq).slice(-3);

    const row = [
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      receiptNo,
      safe_(entry.name),
      safe_(entry.kana),
      entry.phone,
      safe_(entry.email),
      ENTRY_LABEL,
      String(entry.companions),
      String(entry.total),
      entry.member,
      safe_(entry.referrer),
      entry.nextInfo ? '希望する' : '',
      safe_(entry.message),
      '',
    ];
    const range = sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length);
    range.setNumberFormat('@'); // 電話番号の先頭0が消えないよう文字列として保存
    range.setValues([row]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try {
    sendConfirmation_(entry, receiptNo);
    notifyOrganizer_(entry, receiptNo);
  } catch (err) {
    console.error('メール送信に失敗しました', err);
  }
  return json_({ ok: true, receiptNo: receiptNo });
}

// ---------- 内部処理 ----------

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#e7f2e1');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(13, 300);
  }
  return sheet;
}

function status_(sheet) {
  // 以前の「第一部／第二部」形式の行も含め、キャンセル以外の合計人数を数える
  let used = 0;
  const last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, HEADERS.length).getValues().forEach((r) => {
      if (String(r[COL.cancel]).trim()) return;
      if (!String(r[COL.part]).trim()) return;
      used += Number(r[COL.total]) || 1;
    });
  }

  const left = Math.max(0, capacity_() - used);
  // part1・part2 は以前の版のフォーム向け（中身は同じ値）
  return { ok: true, open: !isClosed_(), deadline: CONFIG.DEADLINE, remaining: { all: left, part1: left, part2: left } };
}

// 以前の { part1: 40, part2: 40 } 形式の設定でも動くようにする
function capacity_() {
  const c = CONFIG.CAPACITY;
  return typeof c === 'number' ? c : Math.min.apply(null, Object.keys(c).map((k) => c[k]));
}

function isClosed_() {
  return new Date() > new Date(CONFIG.DEADLINE);
}

function validate_(d) {
  const s = (v, max) => String(v == null ? '' : v).trim().slice(0, max);
  const entry = {
    name: s(d.name, 50),
    kana: s(d.kana, 50),
    phone: s(d.phone, 20),
    email: s(d.email, 100),
    companions: parseInt(d.companions, 10),
    member: s(d.member, 10),
    referrer: s(d.referrer, 50),
    message: s(d.message, 1000),
    nextInfo: d.nextInfo === true,
    companionNames: (Array.isArray(d.companionNames) ? d.companionNames : []).map((n) => s(n, 50)).filter(Boolean),
  };
  if (!entry.name) return { error: 'お名前を入力してください。' };
  if (!entry.kana) return { error: 'ふりがなを入力してください。' };
  const digits = entry.phone.replace(/\D/g, '');
  if (!/^[0-9+\-() ]+$/.test(entry.phone) || digits.length < 10 || digits.length > 15) {
    return { error: '電話番号を正しく入力してください。' };
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(entry.email)) return { error: 'メールアドレスを正しく入力してください。' };
  if (!(entry.companions >= 0 && entry.companions <= 5)) return { error: '同伴者の人数を選んでください。' };
  // お名前のない同伴者の席は確保しない（枠だけ先に取られるのを防ぐ）
  if (entry.companionNames.length < entry.companions) return { error: '一緒に来られる方全員のお名前をご記入ください。' };
  if (['会員', '非会員'].indexOf(entry.member) < 0) return { error: '会員区分を選んでください。' };
  entry.total = entry.companions + 1;
  return { entry: entry };
}

// 「=」などで始まる入力がスプレッドシートの数式として動かないようにする
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function entryLines_(entry) {
  const lines = [
    '■ 日時',
    EVENT.dateLabel,
    EVENT.timeLabel,
  ];
  lines.push(
    '',
    '■ 会場',
    EVENT.venue,
    EVENT.address,
    '地図：' + EVENT.mapUrl,
    '',
    '■ お申し込み内容',
    'お名前：' + entry.name + '（' + entry.kana + '）',
    '電話番号：' + entry.phone,
    '人数：' + entry.total + '名' + (entry.companions ? '（ご本人＋' + entry.companions + '名）' : ''),
    ...(entry.companions ? ['一緒に来られる方：' + entry.companionNames.join('、')] : []),
    '会員区分：' + entry.member,
    'ご紹介者：' + (entry.referrer || '—'),
    '次回のご案内：' + (entry.nextInfo ? '希望する' : '希望しない'),
    'メッセージ：' + (entry.message || '—'),
  );

  const notes = [];
  notes.push('・参加費 1,000円は、当日受付にてお支払いください。');
  if (entry.member === '会員') notes.push('・会員様は、会場協力費 1,000円を追加でお預かりいたします。');
  notes.push('・スリッパをご持参いただけると助かります。');
  if (notes.length) lines.push('', '■ 当日のご案内', ...notes);
  return lines;
}

function sendConfirmation_(entry, receiptNo) {
  const body = [
    entry.name + ' 様',
    '',
    'このたびは「' + EVENT.title + '」にお申し込みいただき、誠にありがとうございます。',
    '以下の内容で受け付けいたしました。',
    '',
    '■ 受付番号：' + receiptNo,
    '',
  ].concat(entryLines_(entry), [
    '',
    'ご都合が悪くなった場合は、このメールに返信してお知らせください。',
    '当日お会いできるのを楽しみにしております。',
    '大切な方と一緒に、ぜひ遊びに来てください。',
    '',
    '――――――――――――――――',
    CONFIG.SENDER_NAME,
  ]);
  if (CONFIG.CONTACT) body.push('お問い合わせ：' + CONFIG.CONTACT);

  const options = { name: CONFIG.SENDER_NAME };
  if (CONFIG.ORGANIZER_EMAIL) options.replyTo = CONFIG.ORGANIZER_EMAIL;
  MailApp.sendEmail(entry.email, '【お申し込み受付】' + EVENT.title + '（11/23）', body.join('\n'), options);
}

function notifyOrganizer_(entry, receiptNo) {
  if (!CONFIG.ORGANIZER_EMAIL) return;
  const status = status_(getSheet_());
  const body = [
    '新しいお申し込みがありました。',
    '',
    '受付番号：' + receiptNo,
    'メール：' + entry.email,
    '',
  ].concat(entryLines_(entry), [
    '',
    '■ 現在の残席',
    '残り ' + status.remaining.all + '席（定員 ' + capacity_() + '名）',
    '',
    '申込一覧：' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  ]);
  MailApp.sendEmail(CONFIG.ORGANIZER_EMAIL, '【新規申込 ' + receiptNo + '】' + entry.name + ' 様（' + entry.total + '名）', body.join('\n'), { name: CONFIG.SENDER_NAME });
}

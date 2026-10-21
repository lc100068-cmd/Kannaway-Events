/**
 * ロバート・ケリー氏と語ろう会 お申し込み受付スクリプト
 *
 * Google スプレッドシートの「拡張機能 → Apps Script」に貼り付けて使います。
 * 設定を変えるときは CONFIG だけを書き換えてください。
 */
const CONFIG = {
  // 定員（同伴者を含む人数）
  CAPACITY: { part1: 40, part2: 40 },
  // 申込締切（日本時間）。config.js の DEADLINE と同じにしてください。
  DEADLINE: '2026-11-22T23:59:59+09:00',
  // 申し込みがあったときに通知を受け取るメールアドレス（空なら通知しない）
  ORGANIZER_EMAIL: '',
  // 確認メールの差出人名
  SENDER_NAME: 'ロバート・ケリー氏と語ろう会 事務局',
  // 確認メールの末尾に載せるお問い合わせ先（空なら載せない）
  CONTACT: '',
  // 申込一覧を書き込むシート名
  SHEET_NAME: '申込一覧',
};

const EVENT = {
  title: 'ロバート・ケリー氏と語ろう会',
  dateLabel: '2026年11月23日（月・祝）',
  venue: 'レンタルスペース GOODLUCK',
  address: '〒703-8203 岡山県岡山市中区国府市場331-3',
  mapUrl: 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent('岡山県岡山市中区国府市場331-3'),
};

const PARTS = {
  both: { label: '第一部・第二部 両方', uses: ['part1', 'part2'] },
  part1: { label: '第一部のみ', uses: ['part1'] },
  part2: { label: '第二部のみ', uses: ['part2'] },
};

const PART_DETAIL = {
  part1: '第一部「ロバート・ケリー氏と語ろう会」　開場 10:00／10:30〜12:00',
  part2: '第二部「「麻」が持つ本当の魅力 キャナウェイ説明会」　開場 13:00／13:30〜16:30',
};

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
    for (const p of PARTS[entry.part].uses) {
      if (entry.total > status.remaining[p]) {
        const left = Math.min(...PARTS[entry.part].uses.map((u) => status.remaining[u]));
        return json_({
          ok: false,
          code: 'full',
          message: left > 0
            ? `申し訳ありません。選択した部の残席は${left}席です。人数か参加する部をご確認ください。`
            : '申し訳ありません。選択した部は満席になりました。',
          remaining: status.remaining,
        });
      }
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
      PARTS[entry.part].label,
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
  const used = { part1: 0, part2: 0 };
  const labelToKey = {};
  Object.keys(PARTS).forEach((k) => { labelToKey[PARTS[k].label] = k; });

  const last = sheet.getLastRow();
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, HEADERS.length).getValues().forEach((r) => {
      if (String(r[COL.cancel]).trim()) return;
      const key = labelToKey[r[COL.part]];
      if (!key) return;
      const n = Number(r[COL.total]) || 1;
      PARTS[key].uses.forEach((p) => { used[p] += n; });
    });
  }

  const remaining = {};
  Object.keys(CONFIG.CAPACITY).forEach((p) => {
    remaining[p] = Math.max(0, CONFIG.CAPACITY[p] - used[p]);
  });
  return { ok: true, open: !isClosed_(), deadline: CONFIG.DEADLINE, remaining: remaining };
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
    part: s(d.part, 10),
    companions: parseInt(d.companions, 10),
    member: s(d.member, 10),
    referrer: s(d.referrer, 50),
    message: s(d.message, 1000),
    nextInfo: d.nextInfo === true,
  };
  if (!entry.name) return { error: 'お名前を入力してください。' };
  if (!entry.kana) return { error: 'ふりがなを入力してください。' };
  const digits = entry.phone.replace(/\D/g, '');
  if (!/^[0-9+\-() ]+$/.test(entry.phone) || digits.length < 10 || digits.length > 15) {
    return { error: '電話番号を正しく入力してください。' };
  }
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(entry.email)) return { error: 'メールアドレスを正しく入力してください。' };
  if (!PARTS[entry.part]) return { error: '参加する部を選んでください。' };
  if (!(entry.companions >= 0 && entry.companions <= 5)) return { error: '同伴者の人数を選んでください。' };
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
  const uses = PARTS[entry.part].uses;
  const lines = [
    '■ 日時',
    EVENT.dateLabel,
  ];
  uses.forEach((p) => lines.push(PART_DETAIL[p]));
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
    '参加する部：' + PARTS[entry.part].label,
    '人数：' + entry.total + '名' + (entry.companions ? '（ご本人＋' + entry.companions + '名）' : ''),
    '会員区分：' + entry.member,
    'ご紹介者：' + (entry.referrer || '—'),
    '次回のご案内：' + (entry.nextInfo ? '希望する' : '希望しない'),
    'メッセージ：' + (entry.message || '—'),
  );

  const notes = [];
  if (uses.indexOf('part1') >= 0) notes.push('・第一部の参加費 1,000円は、当日受付にてお支払いください。');
  if (uses.indexOf('part2') >= 0 && entry.member === '会員') notes.push('・第二部は、会員様のみ会場協力費 1,000円をお預かりいたします。');
  if (entry.part === 'both') notes.push('・第一部・第二部を通してご参加の方は、昼食をご持参ください。');
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
    '第一部：' + status.remaining.part1 + '席',
    '第二部：' + status.remaining.part2 + '席',
    '',
    '申込一覧：' + SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  ]);
  MailApp.sendEmail(CONFIG.ORGANIZER_EMAIL, '【新規申込 ' + receiptNo + '】' + entry.name + ' 様（' + PARTS[entry.part].label + '・' + entry.total + '名）', body.join('\n'), { name: CONFIG.SENDER_NAME });
}

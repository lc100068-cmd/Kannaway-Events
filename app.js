(() => {
  'use strict';

  const C = window.EVENT_CONFIG || {};
  const DEMO = !C.GAS_URL;

  const EVENT_TITLE = 'ロバート・ケリー氏と語ろう会';
  const EVENT_DATE = '20261123';
  const DATE_LABEL = '2026年11月23日（月・祝）';
  const VENUE = 'レンタルスペース GOODLUCK';
  const ADDRESS = '〒703-8203 岡山県岡山市中区国府市場331-3';
  const MAP_URL = 'https://www.google.com/maps/search/?api=1&query=' + encodeURIComponent('岡山県岡山市中区国府市場331-3');

  // uses: 定員を消費する部 / start・end: カレンダー登録用の時刻（開場〜終了）
  const PARTS = {
    both: { label: '第一部・第二部 両方', short: '第一部・第二部', uses: ['part1', 'part2'], start: '10:00', end: '16:30', time: '10:00 開場〜16:30' },
    part1: { label: '第一部のみ', short: '第一部', uses: ['part1'], start: '10:00', end: '12:00', time: '10:00 開場／10:30〜12:00' },
    part2: { label: '第二部のみ', short: '第二部', uses: ['part2'], start: '13:00', end: '16:30', time: '13:00 開場／13:30〜16:30' },
  };

  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  const form = $('#entry-form');
  const views = { 1: form, 2: $('#confirm-view'), 3: $('#done-view') };
  let seatStatus = null; // { open, remaining: { part1, part2 } }
  let pending = null;    // 確認画面に表示中の申込データ

  init();

  function init() {
    if (DEMO) $('#demo-banner').hidden = false;
    $$('.js-deadline').forEach((el) => { el.textContent = formatDeadline(C.DEADLINE); });
    $$('.js-map-link').forEach((el) => { el.href = MAP_URL; });
    if (C.CONTACT) {
      const contact = $('#contact');
      contact.textContent = 'お問い合わせ：' + C.CONTACT;
      contact.hidden = false;
    }

    if (C.DEADLINE && Date.now() > Date.parse(C.DEADLINE)) {
      showClosed('お申し込みの受付期間は終了しました。たくさんのお申し込みありがとうございました。');
      return;
    }

    form.addEventListener('input', onInput);
    form.addEventListener('change', onInput);
    form.addEventListener('submit', onReview);
    $('#back-btn').addEventListener('click', () => showStep(1));
    $('#send-btn').addEventListener('click', onSend);
    $('#ics-btn').addEventListener('click', downloadIcs);

    updateTotal();
    loadSeats();
    watchStickyCta();
  }

  // 上部の申し込みボタンとフォームがどちらも見えていないときだけ、下部固定ボタンを出す
  function watchStickyCta() {
    const sticky = $('#sticky-cta');
    if (!sticky || !('IntersectionObserver' in window)) return;
    const visible = new Set();
    const io = new IntersectionObserver((entries) => {
      entries.forEach((en) => (en.isIntersecting ? visible.add(en.target) : visible.delete(en.target)));
      sticky.classList.toggle('is-hidden', visible.size > 0);
    });
    [$('.date-card'), $('#apply')].forEach((el) => io.observe(el));
  }

  // ---------- 残席 ----------

  async function loadSeats() {
    if (DEMO) {
      applySeats({ ok: true, open: true, remaining: { part1: 4, part2: 25 } });
      return;
    }
    try {
      const res = await fetch(C.GAS_URL + '?action=status', { cache: 'no-store' });
      applySeats(await res.json());
    } catch (e) {
      // 残席が取得できなくても申し込みはできる（最終判定はサーバー側）
      $$('.seat-badge').forEach((el) => { el.textContent = ''; });
    }
  }

  function applySeats(s) {
    if (!s || !s.ok || !s.remaining) return;
    seatStatus = s;
    if (!s.open) {
      showClosed('お申し込みの受付は終了しました。たくさんのお申し込みありがとうございました。');
      return;
    }
    const r = s.remaining;
    if (r.part1 <= 0 && r.part2 <= 0) {
      showClosed('おかげさまで満席となりました。たくさんのお申し込みありがとうございました。');
      return;
    }

    $$('.seat-badge').forEach((el) => {
      const left = r[el.dataset.seat];
      el.textContent = seatText(left, '残り');
      el.classList.toggle('few', left > 0 && left <= 5);
      el.classList.toggle('full', left <= 0);
    });

    Object.keys(PARTS).forEach((key) => {
      const left = partRemaining(key);
      const input = form.querySelector(`input[name="part"][value="${key}"]`);
      const choice = input.closest('.choice');
      const seat = $('.seat', choice);
      input.disabled = left <= 0;
      if (input.disabled) input.checked = false;
      choice.classList.toggle('is-full', left <= 0);
      seat.textContent = seatText(left, '残り');
      seat.classList.toggle('few', left > 0 && left <= 5);
    });
  }

  function partRemaining(key) {
    if (!seatStatus) return Infinity;
    return Math.min(...PARTS[key].uses.map((p) => seatStatus.remaining[p]));
  }

  function seatText(left, prefix) {
    if (left <= 0) return '満席';
    if (left <= 5) return `残りわずか（${prefix}${left}席）`;
    return `${prefix}${left}席`;
  }

  // ---------- 入力 ----------

  function onInput(e) {
    updateTotal();
    updateMemberNotice();
    const name = e.target.name;
    if (name) clearError(name);
    if (name === 'companions') clearError('part'); // 人数を変えたら残席エラーも消す
  }

  function updateTotal() {
    $('#total-count').textContent = String(Number(form.companions.value || 0) + 1);
  }

  function updateMemberNotice() {
    const part = radioValue('part');
    const member = radioValue('member');
    let text = '';
    if (member === '会員' && (part === 'part2' || part === 'both')) {
      text = '第二部は、会員様のみ会場協力費として1,000円をお預かりいたします。';
    } else if (member === '非会員' && (part === 'part2' || part === 'both')) {
      text = '第二部は無料でご参加いただけます。';
    }
    $('#member-notice').textContent = text;
  }

  function radioValue(name) {
    const checked = form.querySelector(`input[name="${name}"]:checked`);
    return checked ? checked.value : '';
  }

  // 全角英数・記号を半角に（電話番号・メールの入力ゆれ対策）
  function toHalfWidth(s) {
    return s
      .replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xfee0))
      .replace(/[ー－―‐−]/g, '-')
      .replace(/　/g, ' ');
  }

  function collect() {
    const v = (name) => (form.elements[name].value || '').trim();
    return {
      name: v('name'),
      kana: v('kana'),
      phone: toHalfWidth(v('phone')),
      email: toHalfWidth(v('email')),
      part: radioValue('part'),
      companions: Number(form.companions.value || 0),
      member: radioValue('member'),
      referrer: v('referrer'),
      message: v('message'),
      nextInfo: form.elements.nextInfo.checked,
      website: v('website'),
    };
  }

  function validate(d) {
    const errs = {};
    if (!d.name) errs.name = 'お名前を入力してください';
    if (!d.kana) errs.kana = 'ふりがなを入力してください';

    const digits = d.phone.replace(/\D/g, '');
    if (!d.phone) errs.phone = '電話番号を入力してください';
    else if (!/^[0-9+\-() ]+$/.test(d.phone) || digits.length < 10 || digits.length > 15) {
      errs.phone = '電話番号を正しく入力してください（例：090-1234-5678）';
    }

    if (!d.email) errs.email = 'メールアドレスを入力してください';
    else if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) errs.email = 'メールアドレスの形式が正しくありません';

    if (!d.part) errs.part = '参加する部を選んでください';
    else {
      const left = partRemaining(d.part);
      if (d.companions + 1 > left) {
        errs.part = `選択した部の残席は${left}席です。人数か参加する部をご確認ください`;
      }
    }

    if (!d.member) errs.member = '会員か非会員かを選んでください';
    return errs;
  }

  function showErrors(errs) {
    let first = null;
    Object.keys(errs).forEach((name) => {
      const el = $(`[data-err="${name}"]`);
      if (el) el.textContent = errs[name];
      setInvalid(name, true);
      if (!first) first = name;
    });
    if (first) {
      const target = form.elements[first];
      const node = target instanceof RadioNodeList ? target[0] : target;
      node.closest('.field').scrollIntoView({ behavior: 'smooth', block: 'center' });
      if (!(target instanceof RadioNodeList)) setTimeout(() => node.focus({ preventScroll: true }), 300);
    }
  }

  function clearError(name) {
    const el = $(`[data-err="${name}"]`);
    if (el) el.textContent = '';
    setInvalid(name, false);
  }

  function setInvalid(name, on) {
    const target = form.elements[name];
    if (!target) return;
    if (target instanceof RadioNodeList) {
      target[0].closest('fieldset').classList.toggle('fieldset-invalid', on);
    } else {
      target.setAttribute('aria-invalid', on ? 'true' : 'false');
    }
  }

  // ---------- 確認 ----------

  function onReview(e) {
    e.preventDefault();
    const data = collect();
    $$('.err', form).forEach((el) => { el.textContent = ''; });
    ['name', 'kana', 'phone', 'email', 'part', 'member'].forEach((n) => setInvalid(n, false));

    const errs = validate(data);
    if (Object.keys(errs).length) {
      showErrors(errs);
      return;
    }
    pending = data;
    renderConfirm(data);
    $('#send-error').hidden = true;
    showStep(2);
  }

  function renderConfirm(d) {
    const rows = [
      ['お名前', d.name],
      ['ふりがな', d.kana],
      ['電話番号', d.phone],
      ['メール', d.email],
      ['参加する部', `${PARTS[d.part].label}\n${PARTS[d.part].time}`],
      ['人数', d.companions ? `${d.companions + 1}名（ご本人＋${d.companions}名）` : '1名（ご本人のみ）'],
      ['会員区分', d.member],
      ['ご紹介者', d.referrer || '—'],
      ['メッセージ', d.message || '—'],
      ['次回のご案内', d.nextInfo ? '希望する' : '希望しない'],
    ];
    const list = $('#confirm-list');
    list.textContent = '';
    rows.forEach(([k, v]) => {
      const row = document.createElement('div');
      const dt = document.createElement('dt');
      const dd = document.createElement('dd');
      dt.textContent = k;
      dd.textContent = v;
      row.append(dt, dd);
      list.append(row);
    });
  }

  // ---------- 送信 ----------

  async function onSend() {
    const btn = $('#send-btn');
    const errBox = $('#send-error');
    btn.disabled = true;
    btn.textContent = '送信中…';
    errBox.hidden = true;

    try {
      let result;
      if (DEMO) {
        await new Promise((r) => setTimeout(r, 700));
        result = { ok: true, receiptNo: 'DEMO-001' };
      } else {
        const res = await fetch(C.GAS_URL, {
          method: 'POST',
          // text/plain にすることで Apps Script へのCORSプリフライトを避ける
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          body: JSON.stringify(pending),
        });
        result = await res.json();
      }

      if (!result.ok) {
        if (result.code === 'closed') {
          showClosed(result.message);
          return;
        }
        if (result.code === 'full') {
          if (result.remaining) applySeats({ ok: true, open: true, remaining: result.remaining });
          showStep(1);
          showErrors({ part: result.message });
          return;
        }
        throw new Error(result.message || '送信できませんでした。');
      }
      showDone(pending, result.receiptNo);
    } catch (err) {
      errBox.textContent = err instanceof TypeError
        ? '通信エラーが発生しました。電波の良い場所で、もう一度「申し込む」を押してください。'
        : err.message;
      errBox.hidden = false;
    } finally {
      btn.disabled = false;
      btn.textContent = 'この内容で申し込む';
    }
  }

  // ---------- 完了 ----------

  function showDone(d, receiptNo) {
    $('#receipt-no').textContent = receiptNo;
    $('#done-email').textContent = d.email;

    const notes = [
      `日時：${DATE_LABEL} ${PARTS[d.part].time}`,
      `会場：${VENUE}（${ADDRESS}）`,
    ];
    if (PARTS[d.part].uses.includes('part1')) notes.push('第一部の参加費 1,000円は、当日受付にてお支払いください。');
    if (PARTS[d.part].uses.includes('part2') && d.member === '会員') notes.push('第二部は、会員様のみ会場協力費 1,000円をお預かりいたします。');
    if (d.part === 'both') notes.push('第一部・第二部を通してご参加の方は、昼食をご持参ください。');

    const ul = $('#done-notes');
    ul.textContent = '';
    notes.forEach((t) => {
      const li = document.createElement('li');
      li.textContent = t;
      ul.append(li);
    });

    $('#gcal-link').href = googleCalendarUrl(d.part);
    form.reset();
    updateTotal();
    showStep(3);
  }

  function calendarDetails(partKey) {
    return `${EVENT_TITLE}（${PARTS[partKey].short}）\n${PARTS[partKey].time}\n会場：${VENUE}\n${ADDRESS}`;
  }

  function googleCalendarUrl(partKey) {
    const p = PARTS[partKey];
    const hm = (t) => t.replace(':', '') + '00';
    const params = new URLSearchParams({
      action: 'TEMPLATE',
      text: `${EVENT_TITLE}（${p.short}）`,
      dates: `${EVENT_DATE}T${hm(p.start)}/${EVENT_DATE}T${hm(p.end)}`,
      ctz: 'Asia/Tokyo',
      location: `${VENUE} ${ADDRESS}`,
      details: calendarDetails(partKey),
    });
    return 'https://calendar.google.com/calendar/render?' + params.toString();
  }

  function downloadIcs() {
    const partKey = (pending && pending.part) || 'both';
    const p = PARTS[partKey];
    // 日本時間 → UTC（-9時間）
    const utc = (t) => {
      const [h, m] = t.split(':').map(Number);
      return `${EVENT_DATE}T${String(h - 9).padStart(2, '0')}${String(m).padStart(2, '0')}00Z`;
    };
    const esc = (s) => s.replace(/\\/g, '\\\\').replace(/[,;]/g, (c) => '\\' + c).replace(/\n/g, '\\n');
    const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}/, '');
    const ics = [
      'BEGIN:VCALENDAR',
      'VERSION:2.0',
      'PRODID:-//Kannaway Events//JA',
      'CALSCALE:GREGORIAN',
      'BEGIN:VEVENT',
      `UID:${Date.now()}-${partKey}@kannaway-events`,
      `DTSTAMP:${stamp}`,
      `DTSTART:${utc(p.start)}`,
      `DTEND:${utc(p.end)}`,
      `SUMMARY:${esc(`${EVENT_TITLE}（${p.short}）`)}`,
      `LOCATION:${esc(`${VENUE} ${ADDRESS}`)}`,
      `DESCRIPTION:${esc(calendarDetails(partKey))}`,
      'END:VEVENT',
      'END:VCALENDAR',
    ].join('\r\n');

    const url = URL.createObjectURL(new Blob([ics], { type: 'text/calendar;charset=utf-8' }));
    const a = document.createElement('a');
    a.href = url;
    a.download = 'robert-kelly-20261123.ics';
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ---------- 画面切り替え ----------

  function showStep(n) {
    Object.entries(views).forEach(([k, el]) => { el.hidden = Number(k) !== n; });
    $$('#steps li').forEach((li) => {
      const s = Number(li.dataset.step);
      li.classList.toggle('is-current', s === n);
      li.classList.toggle('is-done', s < n);
    });
    $('#apply').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function showClosed(message) {
    Object.values(views).forEach((el) => { el.hidden = true; });
    $('#steps').hidden = true;
    $('#closed-message').textContent = message;
    $('#closed-view').hidden = false;
    $$('.cta, .sticky-cta').forEach((el) => { el.hidden = true; });
  }

  function formatDeadline(iso) {
    if (!iso) return '';
    const d = new Date(Date.parse(iso) + 9 * 3600 * 1000); // 日本時間で表示
    const dow = '日月火水木金土'[d.getUTCDay()];
    const hh = String(d.getUTCHours()).padStart(2, '0');
    const mm = String(d.getUTCMinutes()).padStart(2, '0');
    return `${d.getUTCMonth() + 1}月${d.getUTCDate()}日（${dow}）${hh}:${mm}`;
  }
})();

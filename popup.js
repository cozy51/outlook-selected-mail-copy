'use strict';
const status = document.getElementById('status');
const preview = document.getElementById('preview');
const copyBtn = document.getElementById('copy');
const reloadBtn = document.getElementById('reload');
const metadata = document.getElementById('metadata');
const target = document.getElementById('target');
let message = null;
let copiedText = null; // 最後にクリップボードへ書き込んだ内容
function format(m) {
  if (!metadata.checked) return m.body;
  // 取得できなかった項目は見出しも含めて表示しない。
  const header = [
    ['件名', m.subject],
    ['差出人', m.sender],
    ['日時', m.date]
  ].filter(([, value]) => typeof value === 'string' && value.trim())
   .map(([label, value]) => `${label}: ${value.trim()}`);
  return header.length ? [...header, '', m.body].join('\n') : m.body;
}
function render() {
  preview.textContent = message ? format(message) : '対象メールがありません。';
  copyBtn.disabled = !message;
  updateCopiedState();
}
// 表示中の内容がコピー済みかどうかをボタン・プレビューに反映する。
function updateCopiedState() {
  const copied = !!message && copiedText === format(message);
  copyBtn.textContent = copied ? '✓ コピー済み' : 'この1件をコピー';
  copyBtn.classList.toggle('copied', copied);
  preview.classList.toggle('copied', copied);
}
// reason: 'open'（取得直後の自動コピー）/ 'switch'（形式切替）/ 'manual'（ボタン）
async function copyCurrent(reason) {
  const auto = reason !== 'manual';
  if (!message) return;
  const text = format(message);
  try {
    await navigator.clipboard.writeText(text);
    copiedText = text;
    status.textContent = {
      open: '開いた時点でこのメール1件をコピーしました。',
      switch: '切り替え後の内容でコピーし直しました。',
      manual: 'このメール1件をコピーしました。'
    }[reason];
    status.classList.remove('error'); status.classList.add('ok');
  } catch (e) {
    copiedText = null;
    status.textContent = auto
      ? `自動コピーできませんでした（${e.message}）。「この1件をコピー」を押してください。`
      : `コピー失敗: ${e.message}`;
    status.classList.remove('ok'); status.classList.add('error');
  }
  updateCopiedState();
}
async function load() {
  message = null; copiedText = null; copyBtn.disabled = true;
  updateCopiedState();
  status.textContent = '選択中のメールを確認しています…'; status.classList.remove('error', 'ok');
  preview.textContent = '取得中…';
  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab?.id || !/^https:\/\/(outlook\.office\.com|outlook\.office365\.com|outlook\.live\.com|outlook\.cloud\.microsoft)(\/|$)/i.test(tab.url || '')) {
      throw new Error('OutlookのWeb版を開いてください。');
    }
    const result = await chrome.scripting.executeScript({ target: {tabId: tab.id}, func: extractCurrentMail, args: [target.value] });
    const value = result?.[0]?.result;
    if (!value?.ok) throw new Error(value?.error || 'メールを取得できませんでした。');
    message = value;
    render();
    // ボタンを押さなくても、取得できた時点でクリップボードへコピーする。
    await copyCurrent('open');
    if (value.notes?.length && status.classList.contains('ok')) status.textContent += `（${value.notes.join('、')}）`;
  } catch (e) {
    status.textContent = e.message || String(e); status.classList.remove('ok'); status.classList.add('error');
    preview.textContent = 'Outlookで対象のメールを1件開いてから「再取得」を押してください。';
  }
}
metadata.addEventListener('change', () => {
  render();
  // 形式を切り替えたら、切り替え後の内容でコピーし直す。
  if (message) copyCurrent('switch');
  chrome.storage.local.set({ includeMetadata: metadata.checked }).catch(e => {
    status.textContent = `設定を保存できませんでした: ${e.message}`;
    status.classList.add('error');
  });
});
target.addEventListener('change', () => {
  chrome.storage.local.set({ targetMode: target.value }).catch(e => {
    status.textContent = `設定を保存できませんでした: ${e.message}`;
    status.classList.add('error');
  });
  // 対象の決め方が変わるので読み取りからやり直す。
  load();
});
reloadBtn.addEventListener('click', load);
copyBtn.addEventListener('click', () => copyCurrent('manual'));
// 前回のチェック状態を復元してからメールを読み取る。
(async () => {
  try {
    const saved = await chrome.storage.local.get(['includeMetadata', 'targetMode']);
    if (typeof saved.includeMetadata === 'boolean') metadata.checked = saved.includeMetadata;
    // 未設定のときは「最新メールのみ」を初期値とする。
    if (saved.targetMode === 'latest' || saved.targetMode === 'opened') target.value = saved.targetMode;
  } catch (e) {
    status.textContent = `設定の読み込みに失敗しました: ${e.message}`;
    status.classList.add('error');
  }
  await load();
})();

// Outlookの表示DOMは変更されるため、本文を安全に特定できないときはコピーしない。
// mode: 'latest'（最新メールのみ）/ 'opened'（開いているメール1件をそのまま）
function extractCurrentMail(mode) {
  const latestOnly = mode !== 'opened';
  const notes = [];
  const visible = el => {
    if (!el || !(el instanceof HTMLElement)) return false;
    const s = getComputedStyle(el), r = el.getBoundingClientRect();
    return s.visibility !== 'hidden' && s.display !== 'none' && r.width > 10 && r.height > 4;
  };
  const clean = s => (s || '').replace(/\u00a0/g, ' ').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  const one = (root, selectors) => {
    for (const selector of selectors) {
      const elements = [...root.querySelectorAll(selector)].filter(visible);
      if (elements.length) return elements[0];
    }
    return null;
  };
  // チェックボックスによる複数選択は本文の取得とは別なので、読み取りを停止する。
  const checked = [...document.querySelectorAll('[role="checkbox"][aria-checked="true"], input[type="checkbox"]:checked')]
    .filter(visible).filter(el => !!el.closest('[role="option"], [role="row"], [data-convid], [data-itemid]'));
  if (checked.length > 1) return {ok:false,error:'複数メールがチェックされています。対象を1件だけ開いてください。'};
  const bodySelectors = [
    '[data-testid="message-body"]',
    '[data-testid="mail-message-body"]',
    '[aria-label="Message body"]',
    '[aria-label="メッセージ本文"]',
    '[id^="UniqueMessageBody"]',
    '.ReadMsgBody',
    '[data-app-section="MailReadCompose"] [role="document"]'
  ];
  const candidates = bodySelectors.flatMap(sel => [...document.querySelectorAll(sel)].filter(visible));
  const bodies = [...new Set(candidates)].filter(el => !candidates.some(other => other !== el && el.contains(other)));
  const scopeOf = el => el.closest('[data-itemid], [data-message-id], [data-testid="message-container"], [role="article"]') || el.parentElement?.parentElement?.parentElement || document;
  const dateElOf = scope => one(scope, [
    '[data-testid="message-date"]', '[data-testid="sent-time"]',
    'time[datetime]', '[title*="202"]'
  ]);
  // Outlookの表示日時（日本語・英語表記）を数値化する。判定できなければ NaN。
  const toTime = el => {
    if (!el) return NaN;
    const attr = el.getAttribute('datetime');
    if (attr && !isNaN(Date.parse(attr))) return Date.parse(attr);
    const t = `${el.innerText || ''} ${el.title || ''}`;
    let m = t.match(/(\d{4})[\/\-.年](\d{1,2})[\/\-.月](\d{1,2})日?[^\d]*?(\d{1,2}):(\d{2})/);
    if (m) return new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5]).getTime();
    m = t.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?/i);
    if (m) {
      let h = +m[4] % 12; if (/pm/i.test(m[6] || '')) h += 12; else if (!m[6]) h = +m[4];
      return new Date(+m[3], m[1] - 1, +m[2], h, +m[5]).getTime();
    }
    return NaN;
  };
  // 返信・転送で本文に付く過去のやり取り（引用部分）の開始行を探して切り落とす。
  const stripQuoted = text => {
    const lines = text.split('\n');
    const near = (i, re) => lines.slice(i + 1, i + 6).some(l => re.test(l.trim()));
    const headerNext = /^(Sent|Date|To|Subject|送信日時|日時|宛先|件名)\s*[:：]/i;
    const idx = lines.findIndex((raw, i) => {
      const l = raw.trim();
      if (i === 0) return false;
      if (/^-{2,}\s*(Original Message|元のメッセージ|オリジナル\s*メッセージ)\s*-{2,}$/i.test(l)) return true;
      if (/^_{8,}$/.test(l) && near(i, /^(From|差出人|送信者)\s*[:：]/i)) return true;
      if (/^(From|差出人|送信者)\s*[:：]/i.test(l) && near(i, headerNext)) return true;
      if (/^On .+wrote:$/i.test(l)) return true;
      if (/^>/.test(l) && lines.slice(i).filter(x => x.trim()).every(x => /^>/.test(x.trim()))) return true;
      return false;
    });
    if (idx <= 0) return text;
    const head = clean(lines.slice(0, idx).join('\n'));
    if (!head) return text;
    notes.push('過去のやり取りの引用部分を除きました');
    return head;
  };
  if (bodies.length === 0) {
    return {ok:false,error:'メール本文を特定できません。メールを開くか、Outlookの表示形式に合わせた調整が必要です。'};
  }
  let bodyEl = bodies[0];
  if (bodies.length > 1) {
    // 「開いているメール」モードでは1件に絞れないので拒否する。
    if (!latestOnly) return {ok:false,error:'本文が複数見つかりました。スレッド内でコピー対象1件だけを展開してください。'};
    // 「最新メールのみ」モードでは日時が最も新しい1件を選ぶ。日時が読めないものがあれば誤選択を避けて停止。
    const timed = bodies.map(el => ({el, time: toTime(dateElOf(scopeOf(el)))}));
    if (timed.some(x => isNaN(x.time))) {
      return {ok:false,error:'スレッド内のメールの日時を判定できず、最新の1件を特定できません。最新のメールだけを展開してください。'};
    }
    timed.sort((a, b) => b.time - a.time);
    if (timed[0].time === timed[1].time) {
      return {ok:false,error:'同じ日時のメールが複数あり、最新の1件を特定できません。対象1件だけを展開してください。'};
    }
    bodyEl = timed[0].el;
    notes.push(`展開中の${bodies.length}件から最新の1件を選びました`);
  }
  let body = clean(bodyEl.innerText);
  if (!body) return {ok:false,error:'本文が空か、取得できませんでした。'};
  if (latestOnly) body = stripQuoted(body);
  const scope = scopeOf(bodyEl);
  const subjectEl = one(document, [
    '[data-testid="message-subject"]', '[data-testid="mail-subject"]',
    '[data-app-section="MailReadCompose"] [role="heading"]',
    '[data-app-section="MailReadCompose"] h1'
  ]);
  const senderEl = one(scope, [
    '[data-testid="message-sender"]', '[data-testid="sender-name"]',
    '[data-testid="sender-email"]', '[aria-label^="差出人:"]', '[aria-label^="From:"]'
  ]);
  const dateEl = dateElOf(scope);
  return {ok:true,notes,subject:clean(subjectEl?.innerText),sender:clean(senderEl?.innerText || senderEl?.getAttribute('aria-label')),
    date:clean(dateEl?.getAttribute('datetime') || dateEl?.innerText || dateEl?.title),body};
}

'use strict';
const status = document.getElementById('status');
const preview = document.getElementById('preview');
const copyBtn = document.getElementById('copy');
const reloadBtn = document.getElementById('reload');
const metadata = document.getElementById('metadata');
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
    const result = await chrome.scripting.executeScript({ target: {tabId: tab.id}, func: extractCurrentMail });
    const value = result?.[0]?.result;
    if (!value?.ok) throw new Error(value?.error || 'メールを取得できませんでした。');
    message = value;
    render();
    // ボタンを押さなくても、取得できた時点でクリップボードへコピーする。
    await copyCurrent('open');
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
reloadBtn.addEventListener('click', load);
copyBtn.addEventListener('click', () => copyCurrent('manual'));
// 前回のチェック状態を復元してからメールを読み取る。
(async () => {
  try {
    const saved = await chrome.storage.local.get('includeMetadata');
    if (typeof saved.includeMetadata === 'boolean') metadata.checked = saved.includeMetadata;
  } catch (e) {
    status.textContent = `設定の読み込みに失敗しました: ${e.message}`;
    status.classList.add('error');
  }
  await load();
})();

// Outlookの表示DOMは変更されるため、本文を安全に特定できないときはコピーしない。
function extractCurrentMail() {
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
  // スレッド中の複数メッセージが展開されている場合、1件だけに絞れないので拒否。
  if (bodies.length !== 1) {
    return {ok:false,error:bodies.length === 0
      ? 'メール本文を特定できません。メールを開くか、Outlookの表示形式に合わせた調整が必要です。'
      : '本文が複数見つかりました。スレッド内でコピー対象1件だけを展開してください。'};
  }
  const bodyEl = bodies[0];
  const body = clean(bodyEl.innerText);
  if (!body) return {ok:false,error:'本文が空か、取得できませんでした。'};
  const scope = bodyEl.closest('[data-itemid], [data-message-id], [data-testid="message-container"], [role="article"]') || bodyEl.parentElement?.parentElement?.parentElement || document;
  const subjectEl = one(document, [
    '[data-testid="message-subject"]', '[data-testid="mail-subject"]',
    '[data-app-section="MailReadCompose"] [role="heading"]',
    '[data-app-section="MailReadCompose"] h1'
  ]);
  const senderEl = one(scope, [
    '[data-testid="message-sender"]', '[data-testid="sender-name"]',
    '[data-testid="sender-email"]', '[aria-label^="差出人:"]', '[aria-label^="From:"]'
  ]);
  const dateEl = one(scope, [
    '[data-testid="message-date"]', '[data-testid="sent-time"]',
    'time[datetime]', '[title*="202"]'
  ]);
  return {ok:true,subject:clean(subjectEl?.innerText),sender:clean(senderEl?.innerText || senderEl?.getAttribute('aria-label')),
    date:clean(dateEl?.getAttribute('datetime') || dateEl?.innerText || dateEl?.title),body};
}

/**
 * B-side Story ヒアリングフォーム — GAS バックエンド
 * =====================================================
 * 設定手順:
 * 1. このスクリプトを Google Apps Script に貼り付ける
 * 2. スクリプトプロパティに以下を設定:
 *    - KENmail : 受信メールアドレス
 * 3. 「デプロイ」→「新しいデプロイ」→ 種類: ウェブアプリ
 *    - 実行ユーザー: 自分
 *    - アクセス: 全員
 * 4. 発行されたWebアプリURLを bside-form.html の GAS_URL に貼り付ける
 *
 * フォームルート:
 *   - translation  : 商品の無料英語紹介（B-side Story）
 *   - consultation : インバウンド対応のお困りごと相談
 */

// ── CORS プリフライト対応 ──────────────────────────
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── メイン処理 ────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.formType === 'consultation') {
      sendConsultationEmail(data);
    } else {
      sendTranslationEmail(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 翻訳ルート メール送信 ─────────────────────────
function sendTranslationEmail(data) {
  const recipient = PropertiesService.getScriptProperties().getProperty('KENmail');
  if (!recipient) throw new Error('スクリプトプロパティ KENmail が設定されていません');

  const subject = `【B-side Story / 無料翻訳】${data.productName || '（商品名未入力）'} — ヒアリング回答`;

  const dirLabelMap = {
    'A': '風土的背景',
    'B': '歴史的背景',
    'C': '手作業・独自の製法',
    'D': '日常に根付いた文化的背景',
  };

  const deepAnswerLines = (data.directions || []).map(r => {
    const label = dirLabelMap[r] || r;
    const answer = (data.deepAnswers || {})[r.toLowerCase()] || '（未回答）';
    return `  [${r}] ${label}\n  ${answer}`;
  }).join('\n\n');

  const body = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
B-side Story ヒアリング回答（無料翻訳ルート）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ STEP 1 — 商品基本情報
  商品名     : ${data.productName || '—'}
  価格帯     : ${data.priceRange || '—'}
  カテゴリー : ${data.category || '—'}

■ STEP 2 — 商品の背景（複数選択）
  ${(data.directions || []).map(r => `[${r}] ${dirLabelMap[r] || r}`).join('\n  ') || '—'}

■ STEP 3 — レイヤー情報
  製法の特徴   : ${(data.craftFeatures || []).join('、') || '—'}
  環境・風景   : ${(data.environment || []).join('、') || '—'}
  歴史の長さ   : ${data.history || '—'}
  届けたい感情 : ${(data.emotions || []).join('、') || '—'}

■ STEP 4 — 核心の深掘り
${deepAnswerLines || '  （未回答）'}

■ STEP 5 — ビジュアルと最終確認
  旅人へのメッセージ : ${data.message || '—'}
  写真について       : ${data.photo || '—'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
送信日時 : ${data.submittedAt || new Date().toLocaleString('ja-JP')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();

  MailApp.sendEmail({ to: recipient, subject: subject, body: body });
}

// ── 相談ルート メール送信 ─────────────────────────
function sendConsultationEmail(data) {
  const recipient = PropertiesService.getScriptProperties().getProperty('KENmail');
  if (!recipient) throw new Error('スクリプトプロパティ KENmail が設定されていません');

  const subject = `【B-side Story / お困りごと相談】${data.businessName || '（事業者名未入力）'}`;

  const topicLines = (data.topics || []).map(t => `  ・${t}`).join('\n');
  const industryLines = (data.industries || []).map(i => `  ・${i}`).join('\n');

  const body = `
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
B-side Story ヒアリング回答（お困りごと相談ルート）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

■ STEP 1 — 事業者情報
  事業者名・担当者 : ${data.businessName || '—'}

■ STEP 2 — 業種・活動（複数選択）
${industryLines || '  —'}

■ STEP 3 — 活動内容
  ${data.activityDesc || '—'}

■ STEP 4 — 困っているトピック（複数選択）
${topicLines || '  —'}

■ STEP 5 — 詳細・状況
  ${data.details || '—'}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
送信日時 : ${data.submittedAt || new Date().toLocaleString('ja-JP')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
`.trim();

  MailApp.sendEmail({ to: recipient, subject: subject, body: body });
}

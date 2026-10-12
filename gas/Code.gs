/**
 * LUMEA SHIFT 申し込みフォーム受付（Google Apps Script）
 *
 * - order.html から送られた内容をスプレッドシートに記録
 * - お客様へ注文確認メール、管理者へ通知メールを送信
 * - 金額はフォームから受け取らず、ここで計算し直す（改ざん対策）
 */

/* =========================================================
   設定
   ========================================================= */
const SHEET_NAME = '注文';
const ADMIN_EMAIL = 'you@example.com'; // 通知を受け取るアドレスに変更
const SHOP_NAME = 'LUMEA SHIFT';
const SUPPORT_EMAIL = 'support@lumea-beauty.jp';
const TZ = 'Asia/Tokyo';

// order.html の PLANS / PAYMENTS / COUPONS と同じ内容にそろえる
const PLANS = {
  subscription: { label: '定期便', price: 1980, nextPrice: 3480, shipping: 0, nextShipping: 550 },
  single:       { label: '単品購入', price: 3980, nextPrice: null, shipping: 0, nextShipping: null },
};
// express: かんたん決済（フリガナ任意）／ needsPayment: 申し込み後に支払い操作が必要
const PAYMENTS = {
  card:       { label: 'クレジットカード', fee: 0, needsPayment: true },
  paypay:     { label: 'PayPay', fee: 0, needsPayment: true },
  amazonpay:  { label: 'Amazon Pay', fee: 0, express: true },
  rakutenpay: { label: '楽天ペイ', fee: 0, express: true },
  deferred:   { label: '後払い（コンビニ・銀行）', fee: 0 },
  cod:        { label: '代金引換', fee: 330 },
};
const COUPONS = { LINE500: 500 };

const HEADERS = [
  '受付日時', '注文番号', 'プラン', '支払い方法', '申込経路', 'クーポン',
  '商品代金', '割引', '送料', '手数料', '初回合計', '2回目以降',
  '姓', '名', 'セイ', 'メイ', 'メール', '電話番号',
  '郵便番号', '都道府県', '住所1', '住所2', '時間帯', '置き配', 'メール配信', '流入元', 'ステータス',
];

/* =========================================================
   受信
   ========================================================= */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 同時送信で注文番号が重複しないように

    const data = JSON.parse(e.postData.contents);

    // スパム（隠し項目に入力あり）は保存せず、成功したように見せて終了
    if (data.website) return json({ ok: true, orderId: '-' });

    const errors = validate(data);
    if (errors.length) return json({ ok: false, errors: errors });

    const order = buildOrder(data);
    const sheet = getSheet();
    const orderId = createOrderId(sheet);
    const now = new Date();

    sheet.appendRow([
      Utilities.formatDate(now, TZ, 'yyyy/MM/dd HH:mm:ss'),
      orderId,
      order.plan.label,
      order.payment.label,
      order.payment.express ? 'かんたん決済' : 'フォーム',
      order.coupon,
      order.plan.price,
      order.discount,
      order.plan.shipping,
      order.payment.fee,
      order.total,
      order.nextTotal == null ? '' : order.nextTotal,
      data.lastName, data.firstName, data.lastKana, data.firstKana,
      data.email,
      "'" + data.tel,   // 先頭の0が消えないよう文字列として保存
      "'" + data.zip,
      data.pref, data.address1, data.address2,
      data.deliveryTime,
      data.okihai ? '希望する' : '',
      data.newsletter ? '受け取る' : '',
      data.source || '',
      order.payment.needsPayment ? '決済待ち' : '受付',
    ].map(safeCell));

    sendCustomerMail(data, order, orderId);
    sendAdminMail(data, order, orderId);

    // カード決済を組み込む場合は、ここで決済ページのURLを作って返す
    return json({ ok: true, orderId: orderId /*, paymentUrl: '...' */ });
  } catch (err) {
    console.error(err);
    return json({ ok: false, errors: ['server'] });
  } finally {
    lock.releaseLock();
  }
}

// 動作確認用（ブラウザでウェブアプリのURLを開くと表示）
function doGet() {
  return json({ ok: true, message: 'LUMEA SHIFT order endpoint' });
}

/* =========================================================
   検証・計算
   ========================================================= */
function validate(d) {
  const errors = [];
  const payment = PAYMENTS[d.payment];
  const req = ['lastName', 'firstName', 'email', 'tel', 'zip', 'pref', 'address1'];
  // Amazon Pay などはフリガナを返さないため、かんたん決済では任意
  if (!(payment && payment.express)) req.push('lastKana', 'firstKana');
  req.forEach(function (k) { if (!str(d[k])) errors.push(k); });
  if (!PLANS[d.plan]) errors.push('plan');
  if (!payment) errors.push('payment');
  if (d.coupon && !COUPONS[d.coupon]) errors.push('coupon');
  if (d.agree !== true) errors.push('agree');
  if (str(d.email) && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) errors.push('email');
  if (str(d.tel) && !/^0\d{9,10}$/.test(d.tel)) errors.push('tel');
  if (str(d.zip) && !/^\d{7}$/.test(d.zip)) errors.push('zip');
  // 長すぎる入力を弾く
  Object.keys(d).forEach(function (k) {
    if (typeof d[k] === 'string' && d[k].length > 200) errors.push(k);
  });
  return errors;
}

function buildOrder(d) {
  const plan = PLANS[d.plan];
  const payment = PAYMENTS[d.payment];
  const discount = d.coupon ? COUPONS[d.coupon] : 0;
  const total = plan.price - discount + plan.shipping + payment.fee;
  const nextTotal = plan.nextPrice == null ? null : plan.nextPrice + plan.nextShipping + payment.fee;
  return { plan: plan, payment: payment, coupon: d.coupon || '', discount: discount, total: total, nextTotal: nextTotal };
}

/* =========================================================
   シート
   ========================================================= */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#EAE3D9');
  }
  return sheet;
}

// 例：LS260917-0001（日付＋その日の連番）
function createOrderId(sheet) {
  const today = Utilities.formatDate(new Date(), TZ, 'yyMMdd');
  const prefix = 'LS' + today + '-';
  const lastRow = sheet.getLastRow();
  let count = 0;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    ids.forEach(function (r) { if (String(r[0]).indexOf(prefix) === 0) count++; });
  }
  return prefix + String(count + 1).padStart(4, '0');
}

// =, +, -, @ で始まる入力が数式として実行されないようにする
function safeCell(v) {
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

/* =========================================================
   メール
   ========================================================= */
function sendCustomerMail(d, o, orderId) {
  const isSub = d.plan === 'subscription';
  const lines = [
    d.lastName + ' ' + d.firstName + ' 様',
    '',
    SHOP_NAME + ' をお申し込みいただき、ありがとうございます。',
    '以下の内容で受け付けました。',
    '',
    '――――――――――――――――',
    '■ ご注文番号：' + orderId,
    '■ 商品：LUMEA SHIFT 朝夜兼用オールインワン美容液 50mL × 1本',
    '■ プラン：' + o.plan.label,
    '■ お支払い方法：' + o.payment.label,
    o.discount ? '■ クーポン割引：-' + yen(o.discount) + '（' + o.coupon + '）' : null,
    '■ 送料：' + (o.plan.shipping ? yen(o.plan.shipping) : '無料'),
    o.payment.fee ? '■ 代引手数料：' + yen(o.payment.fee) : null,
    '■ ' + (isSub ? '初回' : '') + 'お支払い合計：' + yen(o.total) + '（税込）',
    isSub ? '■ 2回目以降：約1ヶ月ごとに ' + yen(o.nextTotal) + '（税込・商品 ' + yen(o.plan.nextPrice) + ' ＋ 送料 ' + (o.plan.nextShipping ? yen(o.plan.nextShipping) : '無料') + '）' : null,
    '――――――――――――――――',
    '',
    '【お届け先】',
    '〒' + d.zip.slice(0, 3) + '-' + d.zip.slice(3),
    d.pref + d.address1 + (d.address2 ? ' ' + d.address2 : ''),
    'お届け時間帯：' + d.deliveryTime + (d.okihai ? '／置き配希望' : ''),
    '',
    'お申し込み確定後、3営業日以内に発送します。',
    '',
  ];

  if (isSub) {
    lines.push(
      '【定期便について】',
      '・回数の縛りはありません。初回のお届け後に解約することもできます。',
      '・解約・お休みは、次回お届け予定日の5日前までにご連絡ください。',
      ''
    );
  }
  if (o.payment.needsPayment) {
    lines.push('※ ' + o.payment.label + 'でのお支払いが完了すると、ご注文が確定します。', '');
  }

  lines.push(
    '完璧じゃなくて、いい。',
    '1本で、朝も夜も、整えていきましょう。',
    '',
    '――――――――――――――――',
    'LUMEA BEAUTY株式会社',
    SUPPORT_EMAIL + '（平日 10:00〜18:00）',
    '※ このメールは送信専用です。'
  );

  MailApp.sendEmail({
    to: d.email,
    subject: '【' + SHOP_NAME + '】お申し込みありがとうございます（' + orderId + '）',
    body: lines.filter(function (l) { return l !== null; }).join('\n'),
    name: SHOP_NAME,
    replyTo: SUPPORT_EMAIL,
  });
}

function sendAdminMail(d, o, orderId) {
  const body = [
    '新しい申し込みがありました。',
    '',
    '注文番号：' + orderId,
    'プラン：' + o.plan.label,
    '支払い：' + o.payment.label + (o.payment.express ? '（かんたん決済）' : ''),
    '初回合計：' + yen(o.total) + (o.coupon ? '（クーポン ' + o.coupon + '）' : ''),
    'お名前：' + d.lastName + ' ' + d.firstName,
    'メール：' + d.email,
    '電話：' + d.tel,
    '',
    SpreadsheetApp.getActiveSpreadsheet().getUrl(),
  ].join('\n');

  MailApp.sendEmail(ADMIN_EMAIL, '【新規注文】' + orderId + ' ' + o.plan.label, body);
}

/* =========================================================
   ヘルパー
   ========================================================= */
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function str(v) { return typeof v === 'string' && v.trim() !== ''; }
function yen(n) { return '¥' + Number(n).toLocaleString('ja-JP'); }

/* =========================================================
   テスト（エディタで実行 → シートとメールを確認）
   初回実行時に権限の承認画面が出ます
   ========================================================= */
function testPost() {
  const sample = {
    plan: 'subscription', payment: 'deferred', coupon: 'LINE500',
    lastName: 'テスト', firstName: '花子', lastKana: 'テスト', firstKana: 'ハナコ',
    email: ADMIN_EMAIL, tel: '09012345678', zip: '1500001',
    pref: '東京都', address1: '渋谷区神宮前3-10-5', address2: '',
    deliveryTime: '19〜21時', okihai: true, newsletter: false, agree: true, website: '',
    source: 'instagram / reels / test',
  };
  const res = doPost({ postData: { contents: JSON.stringify(sample) } });
  console.log(res.getContent());
}

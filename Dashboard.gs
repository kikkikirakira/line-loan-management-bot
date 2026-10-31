// ======================================================
// 貸し借りダッシュボード（既存 Code.gs に追加して使用）
// ======================================================
//
// 前提:
// - 既存の Code.gs に以下が定義済み
//   getSpreadsheetId_(), SHEET_NAME
//   COL_ID, COL_LENT_AT, COL_REGISTERED_BY, COL_PERSON,
//   COL_AMOUNT, COL_MEMO, COL_STATUS, COL_PAID_AT,
//   COL_PAID_BY, COL_NOTE, COL_TALK_TYPE, COL_TALK_ID,
//   COL_BORROWER_USER_ID
//   getSheet_(), getDataRows_(), resolveBorrowerUserId_()
//
// Script Properties（任意）:
// - DASHBOARD_PIN      : スマホ画面を開くためのPIN。未設定ならPINなし。
// - DASHBOARD_TALK_ID  : 複数グループでBotを使う場合、表示対象のtalkId。
//                         1グループしかなければ自動判定する。
// - CHANNEL_ACCESS_TOKEN: LINE表示名・グループ名の取得に使用。
// ======================================================

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('貸し借りダッシュボード')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}


// ------------------------------------------------------
// 初期設定情報
// ------------------------------------------------------

function getDashboardConfig() {
  const props = PropertiesService.getScriptProperties();

  return {
    requiresPin: Boolean(props.getProperty('DASHBOARD_PIN')),
    appName: '貸し借りダッシュボード'
  };
}


// ------------------------------------------------------
// ダッシュボードデータ取得
// google.script.run から呼ぶため、関数名末尾に _ を付けない
// ------------------------------------------------------

function getDashboardData(pin) {
  verifyDashboardPin_(pin);

  const sheet = getSheet_();
  const rows = getDataRows_(sheet);
  const talkId = resolveDashboardTalkId_(rows);

  const targetRows = rows.filter(row => {
    const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();
    return rowTalkId === talkId;
  });

  if (targetRows.length === 0) {
    throw new Error(
      'このグループに紐づいた貸付データがありません。' +
      'LINEグループで「既存データをこのグループに移行」を実行してください。'
    );
  }

  const talkType =
    String(targetRows[0][COL_TALK_TYPE - 1] || 'group').trim() || 'group';

  const groupName = getTalkName_(talkType, talkId);

  const loans = targetRows
    .filter(row => String(row[COL_ID - 1] || '').trim())
    .map(row => {
      const lenderUserId = String(row[COL_REGISTERED_BY - 1] || '').trim();
      const paidByUserId = String(row[COL_PAID_BY - 1] || '').trim();
      const rawBorrowerName =
        String(row[COL_PERSON - 1] || '').trim() || '未設定';

      let borrowerUserId =
        String(row[COL_BORROWER_USER_ID - 1] || '').trim();

      // M列がまだ空の古い行でも、登録名からLINE IDを解決できれば
      // ダッシュボード上では同一人物として扱う。
      if (!borrowerUserId && typeof resolveBorrowerUserId_ === 'function') {
        borrowerUserId = resolveBorrowerUserId_(talkId, rawBorrowerName);
      }

      const lenderName =
        getLineDisplayName_(talkType, talkId, lenderUserId);

      const borrowerName = borrowerUserId
        ? getLineDisplayName_(talkType, talkId, borrowerUserId)
        : rawBorrowerName;

      return {
        id: String(row[COL_ID - 1] || ''),
        lentAt: serializeDateValue_(row[COL_LENT_AT - 1]),
        lenderUserId: lenderUserId,
        lenderName: lenderName,
        borrowerUserId: borrowerUserId,
        borrowerName: borrowerName,
        rawBorrowerName: rawBorrowerName,
        amount: Number(row[COL_AMOUNT - 1]) || 0,
        memo: String(row[COL_MEMO - 1] || '').trim(),
        status: String(row[COL_STATUS - 1] || '').trim() || '未返済',
        paidAt: serializeDateValue_(row[COL_PAID_AT - 1]),
        paidByName: paidByUserId === 'WEB_APP'
          ? 'Webアプリ'
          : paidByUserId
            ? getLineDisplayName_(talkType, talkId, paidByUserId)
            : '',
        note: String(row[COL_NOTE - 1] || '').trim()
      };
    });

  loans.sort((a, b) => {
    return parseDateForSort_(b.lentAt) - parseDateForSort_(a.lentAt);
  });

  const summary = buildDashboardSummary_(loans);
  const members = buildMemberStats_(loans);

  return {
    generatedAt: Utilities.formatDate(
      new Date(),
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm:ss'
    ),
    groupName: groupName,
    summary: summary,
    members: members,
    loans: loans
  };
}


// ------------------------------------------------------
// Webアプリから返済済みに更新
// ------------------------------------------------------
//
// google.script.run から呼ばれる公開関数。
// PINを再検証し、現在のダッシュボード対象グループに属する
// 貸付IDだけを更新する。
//
// WebアプリからはLINE User IDを取得できないため、
// 返済報告者LINE IDには "WEB_APP" を記録する。
// ------------------------------------------------------

function markLoanAsPaidFromDashboard(pin, loanId) {
  verifyDashboardPin_(pin);

  const targetLoanId = String(loanId || '').trim();

  if (!targetLoanId) {
    throw new Error('貸付IDがありません。');
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const sheet = getSheet_();
    const rows = getDataRows_(sheet);
    const talkId = resolveDashboardTalkId_(rows);
    const now = Utilities.formatDate(
      new Date(),
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm:ss'
    );

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowLoanId = String(row[COL_ID - 1] || '').trim();
      const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();

      if (
        rowLoanId !== targetLoanId ||
        rowTalkId !== talkId
      ) {
        continue;
      }

      const status = String(row[COL_STATUS - 1] || '').trim();

      if (status === '返済済み') {
        return {
          ok: true,
          alreadyPaid: true,
          loanId: targetLoanId,
          message: 'この貸付はすでに返済済みです。'
        };
      }

      const rowNumber = i + 2;

      sheet
        .getRange(rowNumber, COL_STATUS)
        .setValue('返済済み');

      sheet
        .getRange(rowNumber, COL_PAID_AT)
        .setValue(now);

      sheet
        .getRange(rowNumber, COL_PAID_BY)
        .setValue('WEB_APP');

      const currentNote =
        String(row[COL_NOTE - 1] || '').trim();

      const webNote =
        'Webアプリから返済更新 ' + now;

      const nextNote = currentNote
        ? currentNote + ' / ' + webNote
        : webNote;

      sheet
        .getRange(rowNumber, COL_NOTE)
        .setValue(nextNote);

      return {
        ok: true,
        alreadyPaid: false,
        loanId: targetLoanId,
        paidAt: now,
        message: '返済済みに更新しました。'
      };
    }

    throw new Error(
      'このグループ内に指定された貸付が見つかりませんでした。'
    );

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // ロック取得前に失敗した場合は何もしない
    }
  }
}


// ------------------------------------------------------
// PIN認証
// ------------------------------------------------------

function verifyDashboardPin_(pin) {
  const expected =
    PropertiesService
      .getScriptProperties()
      .getProperty('DASHBOARD_PIN');

  if (!expected) {
    return;
  }

  if (String(pin || '') !== String(expected)) {
    throw new Error('PINが違います。');
  }
}


// ------------------------------------------------------
// 表示対象グループを決定
// ------------------------------------------------------

function resolveDashboardTalkId_(rows) {
  const configured =
    PropertiesService
      .getScriptProperties()
      .getProperty('DASHBOARD_TALK_ID');

  if (configured) {
    return configured;
  }

  const groupTalkIds = new Set();

  rows.forEach(row => {
    const type = String(row[COL_TALK_TYPE - 1] || '').trim();
    const id = String(row[COL_TALK_ID - 1] || '').trim();

    if (type === 'group' && id) {
      groupTalkIds.add(id);
    }
  });

  if (groupTalkIds.size === 1) {
    return Array.from(groupTalkIds)[0];
  }

  if (groupTalkIds.size > 1) {
    throw new Error(
      '複数のLINEグループのデータがあります。' +
      'スクリプトプロパティ DASHBOARD_TALK_ID に表示対象のトークIDを設定してください。'
    );
  }

  const anyTalkIds = new Set();

  rows.forEach(row => {
    const id = String(row[COL_TALK_ID - 1] || '').trim();

    if (id) {
      anyTalkIds.add(id);
    }
  });

  if (anyTalkIds.size === 1) {
    return Array.from(anyTalkIds)[0];
  }

  throw new Error(
    'トークIDを特定できません。' +
    'まず既存データをLINEグループへ移行してください。'
  );
}


// ------------------------------------------------------
// 集計
// ------------------------------------------------------

function buildDashboardSummary_(loans) {
  let total = 0;
  let outstanding = 0;
  let repaid = 0;
  let unpaidCount = 0;
  let repaidCount = 0;

  loans.forEach(loan => {
    total += loan.amount;

    if (loan.status === '返済済み') {
      repaid += loan.amount;
      repaidCount++;
    } else {
      outstanding += loan.amount;
      unpaidCount++;
    }
  });

  const repaymentRate =
    total > 0
      ? Math.round((repaid / total) * 1000) / 10
      : 0;

  return {
    total: total,
    outstanding: outstanding,
    repaid: repaid,
    unpaidCount: unpaidCount,
    repaidCount: repaidCount,
    totalCount: loans.length,
    repaymentRate: repaymentRate
  };
}


// ------------------------------------------------------
// メンバー別集計
//
// LINE User IDがある場合は名前ではなくLINE User IDをキーにする。
// そのため表示名や入力名が変わっても同じ人物のカードへ集約される。
// ------------------------------------------------------

function buildMemberStats_(loans) {
  const map = {};

  function ensure(userId, name) {
    const key = userId
      ? 'user:' + userId
      : 'name:' + normalizePersonKey_(name);

    if (!map[key]) {
      map[key] = {
        userId: userId || '',
        name: name || '未設定',
        lentTotal: 0,
        borrowedTotal: 0,
        receivable: 0,
        payable: 0,
        repaidBorrowing: 0,
        lentCount: 0,
        borrowedCount: 0
      };
    } else if (userId && name) {
      // LINE表示名が変わった場合は現在の表示名を優先
      map[key].name = name;
    }

    return map[key];
  }

  loans.forEach(loan => {
    const lender = ensure(
      loan.lenderUserId,
      loan.lenderName
    );

    const borrower = ensure(
      loan.borrowerUserId,
      loan.borrowerName
    );

    lender.lentTotal += loan.amount;
    lender.lentCount++;

    borrower.borrowedTotal += loan.amount;
    borrower.borrowedCount++;

    if (loan.status === '返済済み') {
      borrower.repaidBorrowing += loan.amount;
    } else {
      lender.receivable += loan.amount;
      borrower.payable += loan.amount;
    }
  });

  return Object.keys(map)
    .map(key => {
      const item = map[key];
      item.net = item.receivable - item.payable;
      return item;
    })
    .sort((a, b) => {
      const aOpen = a.receivable + a.payable;
      const bOpen = b.receivable + b.payable;

      if (bOpen !== aOpen) {
        return bOpen - aOpen;
      }

      return a.name.localeCompare(b.name, 'ja');
    });
}


function normalizePersonKey_(name) {
  return String(name || '未設定')
    .trim()
    .replace(/\s+/g, ' ')
    .toLocaleLowerCase();
}


// ------------------------------------------------------
// LINEプロフィール・グループ名
// ------------------------------------------------------

function getDashboardChannelAccessToken_() {
  return PropertiesService
    .getScriptProperties()
    .getProperty('CHANNEL_ACCESS_TOKEN') || '';
}


function getTalkName_(talkType, talkId) {
  if (talkType !== 'group') {
    return '貸し借りグループ';
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'group_name_' + talkId;
  const cached = cache.get(cacheKey);

  if (cached) {
    return cached;
  }

  const token = getDashboardChannelAccessToken_();

  if (!token) {
    return '貸し借りグループ';
  }

  try {
    const url =
      'https://api.line.me/v2/bot/group/' +
      encodeURIComponent(talkId) +
      '/summary';

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const name = data.groupName || '貸し借りグループ';

      cache.put(cacheKey, name, 21600);

      return name;
    }
  } catch (err) {
    console.error('getTalkName error: ' + err);
  }

  return '貸し借りグループ';
}


function getLineDisplayName_(talkType, talkId, userId) {
  if (!userId) {
    return '不明';
  }

  const cache = CacheService.getScriptCache();
  const cacheKey =
    'display_name_' +
    talkType +
    '_' +
    talkId +
    '_' +
    userId;

  const cached = cache.get(cacheKey);

  if (cached) {
    return cached;
  }

  const token = getDashboardChannelAccessToken_();

  if (!token) {
    return fallbackMemberName_(userId);
  }

  let url = '';

  if (talkType === 'group') {
    url =
      'https://api.line.me/v2/bot/group/' +
      encodeURIComponent(talkId) +
      '/member/' +
      encodeURIComponent(userId);
  } else if (talkType === 'room') {
    url =
      'https://api.line.me/v2/bot/room/' +
      encodeURIComponent(talkId) +
      '/member/' +
      encodeURIComponent(userId);
  } else {
    url =
      'https://api.line.me/v2/bot/profile/' +
      encodeURIComponent(userId);
  }

  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      const displayName =
        String(data.displayName || '').trim();

      if (displayName) {
        cache.put(cacheKey, displayName, 21600);
        return displayName;
      }
    }
  } catch (err) {
    console.error('getLineDisplayName error: ' + err);
  }

  return fallbackMemberName_(userId);
}


function fallbackMemberName_(userId) {
  const value = String(userId || '');

  if (value.length >= 4) {
    return 'メンバー ' + value.slice(-4);
  }

  return 'メンバー';
}


// ------------------------------------------------------
// 日付シリアライズ
// Apps Script → ブラウザへ Date を直接返さないため文字列化
// ------------------------------------------------------

function serializeDateValue_(value) {
  if (!value) {
    return '';
  }

  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return Utilities.formatDate(
      value,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm:ss'
    );
  }

  return String(value);
}


function parseDateForSort_(value) {
  if (!value) {
    return 0;
  }

  const normalized =
    String(value)
      .replace(/\//g, '-')
      .replace(' ', 'T');

  const time = new Date(normalized).getTime();

  return isNaN(time) ? 0 : time;
}

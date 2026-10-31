// ======================================================
// LINE 貸付・返済管理Bot
// LINE IDベースのメンバー管理対応版
// ======================================================

const SHEET_NAME = '返済管理';

function getSpreadsheetId_() {
  const spreadsheetId = PropertiesService
    .getScriptProperties()
    .getProperty('SPREADSHEET_ID');

  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID が設定されていません。');
  }

  return spreadsheetId;
}
const MEMBER_SHEET_NAME = 'メンバー';

// 返済管理シート
const COL_ID = 1;              // A: 貸付ID
const COL_LENT_AT = 2;         // B: 登録日時
const COL_REGISTERED_BY = 3;   // C: 登録者LINE ID（貸した人）
const COL_PERSON = 4;          // D: 相手（入力時の名前。履歴として残す）
const COL_AMOUNT = 5;          // E: 金額
const COL_MEMO = 6;            // F: 内容
const COL_STATUS = 7;          // G: 返済状況
const COL_PAID_AT = 8;         // H: 返済日時
const COL_PAID_BY = 9;         // I: 返済報告者LINE ID
const COL_NOTE = 10;           // J: 備考
const COL_TALK_TYPE = 11;      // K: トーク種別
const COL_TALK_ID = 12;        // L: トークID
const COL_BORROWER_USER_ID = 13; // M: 相手LINE ID
const COL_LAST_REMINDER_AT = 14; // N: 最終督促通知日時
const COL_REMINDER_COUNT = 15;   // O: 督促回数

// メンバーシート
const MEMBER_COL_TALK_ID = 1;     // A: トークID
const MEMBER_COL_USER_ID = 2;     // B: LINE User ID
const MEMBER_COL_ALIAS = 3;       // C: 登録名
const MEMBER_COL_ALIAS_KEY = 4;   // D: 正規化済み登録名
const MEMBER_COL_REGISTERED_AT = 5; // E: 登録日時


// ======================================================
// 初期設定
// ======================================================

function setupSheet() {
  const sheet = getSheet_();

  const headers = [
    '貸付ID',
    '登録日時',
    '登録者LINE ID',
    '相手',
    '金額',
    '内容',
    '返済状況',
    '返済日時',
    '返済報告者LINE ID',
    '備考',
    'トーク種別',
    'トークID',
    '相手LINE ID',
    '最終督促通知日時',
    '督促回数'
  ];

  const currentHeaders =
    sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (currentHeaders[i] === '') {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }

  sheet.setFrozenRows(1);
  setupMemberSheet_();
}


function setupMemberSheet_() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  let sheet = ss.getSheetByName(MEMBER_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(MEMBER_SHEET_NAME);
  }

  const headers = [
    'トークID',
    'LINE User ID',
    '登録名',
    '登録名キー',
    '登録日時'
  ];

  const current =
    sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (current[i] === '') {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }

  sheet.setFrozenRows(1);
  return sheet;
}


// ======================================================
// LINE Webhook
// ======================================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput('OK');
    }

    const body = JSON.parse(e.postData.contents);
    const events = body.events || [];

    events.forEach(event => {
      if (event.type !== 'message') return;
      if (!event.message || event.message.type !== 'text') return;

      const text = normalizeText_(event.message.text);
      const source = event.source || {};
      const sourceType = source.type || '';
      const userId = source.userId || '';
      const talkId =
        source.groupId ||
        source.roomId ||
        userId ||
        '';
      const replyToken = event.replyToken || '';

      const replyText = handleText_(
        text,
        userId,
        sourceType,
        talkId
      );

      if (replyToken && replyText) {
        replyToLine_(replyToken, replyText);
      }
    });

    return ContentService.createTextOutput('OK');

  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput('ERROR');
  }
}


// ======================================================
// コマンド振り分け
// ======================================================

function handleText_(text, userId, sourceType, talkId) {
  const sheet = getSheet_();
  setupSheet();

  if (text === '既存データをこのグループに移行') {
    return migrateExistingRowsToCurrentGroup_(
      sheet,
      userId,
      sourceType,
      talkId
    );
  }

  if (text === '名前一覧') {
    return listMemberAliases_(sourceType, talkId);
  }

  if (text === '名前登録') {
    return [
      '登録する名前を指定してください。',
      '',
      '例:',
      '名前登録 田中'
    ].join('\n');
  }

  if (text.startsWith('名前登録 ')) {
    const alias = text.slice('名前登録 '.length).trim();

    return registerMemberAlias_(
      alias,
      userId,
      sourceType,
      talkId
    );
  }

  if (text === '使い方') {
    return helpText_();
  }

  if (text === '未返済' || text === '一覧') {
    return listUnpaid_(sheet, talkId);
  }

  if (text.startsWith('貸した ')) {
    return registerLoan_(
      sheet,
      text,
      userId,
      sourceType,
      talkId
    );
  }

  if (text.startsWith('返した ')) {
    return markAsPaid_(
      sheet,
      text,
      userId,
      talkId
    );
  }

  // グループの普通の会話には反応しない
  if (sourceType === 'group' || sourceType === 'room') {
    return null;
  }

  return [
    '形式がわかりませんでした。',
    '',
    '例:',
    '貸した 田中 3000 昼ごはん',
    '返した 田中 3000',
    '名前登録 田中',
    '未返済',
    '',
    '使い方を確認する場合:',
    '使い方'
  ].join('\n');
}


// ======================================================
// 名前登録 / LINE ID紐付け
// ======================================================

function registerMemberAlias_(alias, userId, sourceType, talkId) {
  if (sourceType !== 'group' && sourceType !== 'room') {
    return '名前登録は貸し借りを管理するLINEグループ内で実行してください。';
  }

  if (!userId || !talkId) {
    return 'LINEユーザー情報を取得できませんでした。';
  }

  const cleanAlias = String(alias || '').trim();
  const aliasKey = normalizeAliasKey_(cleanAlias);

  if (!aliasKey) {
    return '登録する名前を入力してください。';
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const memberSheet = setupMemberSheet_();
    const rows = getMemberRows_(memberSheet);

    // 同じグループ内で同じ名前が別人へ登録済みなら拒否
    for (let i = 0; i < rows.length; i++) {
      const rowTalkId = String(rows[i][MEMBER_COL_TALK_ID - 1] || '').trim();
      const rowUserId = String(rows[i][MEMBER_COL_USER_ID - 1] || '').trim();
      const rowAliasKey = String(rows[i][MEMBER_COL_ALIAS_KEY - 1] || '').trim();

      if (
        rowTalkId === talkId &&
        rowAliasKey === aliasKey &&
        rowUserId !== userId
      ) {
        return [
          `「${cleanAlias}」は別のメンバーに登録済みです。`,
          '別の名前を使用してください。'
        ].join('\n');
      }
    }

    let alreadyRegistered = false;

    for (let i = 0; i < rows.length; i++) {
      const rowTalkId = String(rows[i][MEMBER_COL_TALK_ID - 1] || '').trim();
      const rowUserId = String(rows[i][MEMBER_COL_USER_ID - 1] || '').trim();
      const rowAliasKey = String(rows[i][MEMBER_COL_ALIAS_KEY - 1] || '').trim();

      if (
        rowTalkId === talkId &&
        rowUserId === userId &&
        rowAliasKey === aliasKey
      ) {
        alreadyRegistered = true;
        break;
      }
    }

    if (!alreadyRegistered) {
      memberSheet.appendRow([
        talkId,
        userId,
        cleanAlias,
        aliasKey,
        nowString_()
      ]);
    }

    const updatedCount = backfillBorrowerIdsForUser_(
      talkId,
      userId
    );

    const displayName =
      getMemberDisplayNameForBot_(sourceType, talkId, userId);

    return [
      alreadyRegistered ? '名前はすでに登録済みです。' : '名前を登録しました。',
      '',
      `LINE表示名: ${displayName}`,
      `登録名: ${cleanAlias}`,
      `既存データ更新: ${updatedCount}件`,
      '',
      '過去に別の名前でも記録されている場合は、',
      'その名前でも「名前登録 ○○」を送ってください。'
    ].join('\n');

  } catch (err) {
    console.error(err);
    return '名前登録中にエラーが発生しました。';

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // 何もしない
    }
  }
}


function listMemberAliases_(sourceType, talkId) {
  if (sourceType !== 'group' && sourceType !== 'room') {
    return '名前一覧はLINEグループ内で確認してください。';
  }

  const memberSheet = setupMemberSheet_();
  const rows = getMemberRows_(memberSheet);
  const byUser = {};

  rows.forEach(row => {
    const rowTalkId = String(row[MEMBER_COL_TALK_ID - 1] || '').trim();
    const userId = String(row[MEMBER_COL_USER_ID - 1] || '').trim();
    const alias = String(row[MEMBER_COL_ALIAS - 1] || '').trim();

    if (rowTalkId !== talkId || !userId || !alias) return;

    if (!byUser[userId]) {
      byUser[userId] = [];
    }

    if (!byUser[userId].includes(alias)) {
      byUser[userId].push(alias);
    }
  });

  const userIds = Object.keys(byUser);

  if (userIds.length === 0) {
    return [
      'まだ名前登録されているメンバーはいません。',
      '',
      '各メンバーが次のように送ってください。',
      '名前登録 田中'
    ].join('\n');
  }

  const lines = userIds.map(userId => {
    const displayName =
      getMemberDisplayNameForBot_(sourceType, talkId, userId);

    return `${displayName}: ${byUser[userId].join(', ')}`;
  });

  return [
    '登録済みの名前:',
    '',
    lines.join('\n')
  ].join('\n');
}


function backfillBorrowerIdsForUser_(talkId, userId) {
  const sheet = getSheet_();
  const values = getDataRows_(sheet);
  const aliases = getAliasesForUser_(talkId, userId);
  const aliasKeys = new Set(aliases.map(normalizeAliasKey_));
  let count = 0;

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();
    const borrowerUserId = String(row[COL_BORROWER_USER_ID - 1] || '').trim();
    const personKey = normalizeAliasKey_(row[COL_PERSON - 1]);

    if (
      rowTalkId === talkId &&
      !borrowerUserId &&
      aliasKeys.has(personKey)
    ) {
      sheet
        .getRange(i + 2, COL_BORROWER_USER_ID)
        .setValue(userId);
      count++;
    }
  }

  return count;
}


function resolveBorrowerUserId_(talkId, person) {
  const aliasKey = normalizeAliasKey_(person);
  if (!talkId || !aliasKey) return '';

  const memberSheet = setupMemberSheet_();
  const rows = getMemberRows_(memberSheet);

  for (let i = 0; i < rows.length; i++) {
    const rowTalkId = String(rows[i][MEMBER_COL_TALK_ID - 1] || '').trim();
    const userId = String(rows[i][MEMBER_COL_USER_ID - 1] || '').trim();
    const rowAliasKey = String(rows[i][MEMBER_COL_ALIAS_KEY - 1] || '').trim();

    if (rowTalkId === talkId && rowAliasKey === aliasKey) {
      return userId;
    }
  }

  return '';
}


function getAliasesForUser_(talkId, userId) {
  const memberSheet = setupMemberSheet_();
  const rows = getMemberRows_(memberSheet);
  const aliases = [];

  rows.forEach(row => {
    const rowTalkId = String(row[MEMBER_COL_TALK_ID - 1] || '').trim();
    const rowUserId = String(row[MEMBER_COL_USER_ID - 1] || '').trim();
    const alias = String(row[MEMBER_COL_ALIAS - 1] || '').trim();

    if (
      rowTalkId === talkId &&
      rowUserId === userId &&
      alias &&
      !aliases.includes(alias)
    ) {
      aliases.push(alias);
    }
  });

  return aliases;
}


function getMemberRows_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return [];

  return sheet
    .getRange(2, 1, lastRow - 1, MEMBER_COL_REGISTERED_AT)
    .getValues();
}


function normalizeAliasKey_(value) {
  return String(value || '')
    .trim()
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .toLocaleLowerCase();
}


// ======================================================
// 貸付登録
// ======================================================

function registerLoan_(sheet, text, userId, sourceType, talkId) {
  const match = text.match(
    /^貸した\s+(.+?)\s+([0-9,]+)\s*円?(?:\s+(.+))?$/
  );

  if (!match) {
    return [
      '貸付登録の形式が違います。',
      '',
      '例:',
      '貸した 田中 3000 昼ごはん'
    ].join('\n');
  }

  const person = match[1].trim();
  const amount = parseAmount_(match[2]);
  const memo = match[3] ? match[3].trim() : '';

  if (!person || !amount || amount <= 0) {
    return '相手または金額を正しく読み取れませんでした。';
  }

  if (!talkId) {
    return 'トーク情報を取得できませんでした。';
  }

  const borrowerUserId = resolveBorrowerUserId_(talkId, person);
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const id = createLoanId_();
    const now = nowString_();

    sheet.appendRow([
      id,
      now,
      userId,
      person,
      amount,
      memo,
      '未返済',
      '',
      '',
      '',
      sourceType,
      talkId,
      borrowerUserId,
      '',
      0
    ]);

    const lines = [
      '貸付を登録しました。',
      '',
      `貸付ID: ${id}`,
      `相手: ${person}`,
      `金額: ${amount}円`,
      memo ? `内容: ${memo}` : ''
    ].filter(Boolean);

    if (!borrowerUserId && (sourceType === 'group' || sourceType === 'room')) {
      lines.push(
        '',
        `※「${person}」はLINE IDと未紐付けです。`,
        `本人が「名前登録 ${person}」と送ると、過去分を含めて自動統合されます。`
      );
    }

    lines.push(
      '',
      '返済されたら、次のどちらかで送ってください。',
      `返した ${person} ${amount}`,
      `返した ${id}`
    );

    return lines.join('\n');

  } catch (err) {
    console.error(err);
    return '貸付登録中にエラーが発生しました。';

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // 何もしない
    }
  }
}


// ======================================================
// 返済処理
// ======================================================

function markAsPaid_(sheet, text, userId, talkId) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(5000);

    const byIdMatch =
      text.match(/^返した\s+(L[0-9]{18})$/);

    if (byIdMatch) {
      return markAsPaidById_(
        sheet,
        byIdMatch[1],
        userId,
        talkId
      );
    }

    const byPersonAmountMatch =
      text.match(/^返した\s+(.+?)\s+([0-9,]+)\s*円?$/);

    if (!byPersonAmountMatch) {
      return [
        '返済登録の形式が違います。',
        '',
        '例:',
        '返した 田中 3000',
        '返した L202605010201231234'
      ].join('\n');
    }

    const person = byPersonAmountMatch[1].trim();
    const amount = parseAmount_(byPersonAmountMatch[2]);

    return markAsPaidByPersonAndAmount_(
      sheet,
      person,
      amount,
      userId,
      talkId
    );

  } catch (err) {
    console.error(err);
    return '返済処理中にエラーが発生しました。';

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // 何もしない
    }
  }
}


function markAsPaidById_(sheet, targetId, userId, talkId) {
  const values = getDataRows_(sheet);
  const now = nowString_();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowNumber = i + 2;
    const rowId = String(row[COL_ID - 1] || '').trim();
    const status = String(row[COL_STATUS - 1] || '').trim();
    const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();

    if (rowId === targetId && rowTalkId === talkId) {
      if (status === '返済済み') {
        return [
          'この貸付IDはすでに返済済みです。',
          `貸付ID: ${targetId}`
        ].join('\n');
      }

      sheet.getRange(rowNumber, COL_STATUS).setValue('返済済み');
      sheet.getRange(rowNumber, COL_PAID_AT).setValue(now);
      sheet.getRange(rowNumber, COL_PAID_BY).setValue(userId);

      return [
        '返済済みに更新しました。',
        '',
        `貸付ID: ${targetId}`,
        `返済日時: ${now}`
      ].join('\n');
    }
  }

  return [
    'このトーク内に指定された貸付IDが見つかりませんでした。',
    `貸付ID: ${targetId}`
  ].join('\n');
}


function markAsPaidByPersonAndAmount_(sheet, person, amount, userId, talkId) {
  const values = getDataRows_(sheet);
  const matches = [];
  const targetBorrowerUserId = resolveBorrowerUserId_(talkId, person);
  const targetAliasKey = normalizeAliasKey_(person);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowPerson = String(row[COL_PERSON - 1] || '').trim();
    const rowAmount = Number(row[COL_AMOUNT - 1]) || 0;
    const status = String(row[COL_STATUS - 1] || '').trim();
    const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();
    let rowBorrowerUserId = String(row[COL_BORROWER_USER_ID - 1] || '').trim();

    if (rowTalkId !== talkId || rowAmount !== amount || status !== '未返済') {
      continue;
    }

    // 古い行でM列が空なら、登録名から解決してその場で補完
    if (!rowBorrowerUserId) {
      rowBorrowerUserId = resolveBorrowerUserId_(talkId, rowPerson);

      if (rowBorrowerUserId) {
        sheet
          .getRange(i + 2, COL_BORROWER_USER_ID)
          .setValue(rowBorrowerUserId);
      }
    }

    let isMatch = false;

    if (targetBorrowerUserId) {
      // LINE IDが分かる場合はLINE IDで判定
      isMatch = rowBorrowerUserId === targetBorrowerUserId;
    } else {
      // 未登録の名前の場合のみ従来どおり名前で判定
      isMatch = normalizeAliasKey_(rowPerson) === targetAliasKey;
    }

    if (isMatch) {
      matches.push({
        rowNumber: i + 2,
        row: row
      });
    }
  }

  if (matches.length === 0) {
    return [
      '該当する未返済データが見つかりませんでした。',
      '',
      `相手: ${person}`,
      `金額: ${amount}円`
    ].join('\n');
  }

  if (matches.length >= 2) {
    const candidates = matches
      .slice(0, 5)
      .map(item => {
        const row = item.row;
        const id = row[COL_ID - 1];
        const memo = row[COL_MEMO - 1];
        const lentAt = row[COL_LENT_AT - 1];

        return `${id} / ${lentAt} / ${memo || '内容なし'}`;
      })
      .join('\n');

    return [
      '同じ相手・同じ金額の未返済データが複数あります。',
      'どれを返済済みにするか判断できないため、貸付IDで指定してください。',
      '',
      '候補:',
      candidates,
      '',
      '例:',
      `返した ${matches[0].row[COL_ID - 1]}`
    ].join('\n');
  }

  const target = matches[0];
  const now = nowString_();

  sheet.getRange(target.rowNumber, COL_STATUS).setValue('返済済み');
  sheet.getRange(target.rowNumber, COL_PAID_AT).setValue(now);
  sheet.getRange(target.rowNumber, COL_PAID_BY).setValue(userId);

  const row = target.row;
  const id = row[COL_ID - 1];
  const memo = row[COL_MEMO - 1];

  return [
    '返済済みに更新しました。',
    '',
    `貸付ID: ${id}`,
    `相手: ${person}`,
    `金額: ${amount}円`,
    memo ? `内容: ${memo}` : '',
    `返済日時: ${now}`
  ].filter(Boolean).join('\n');
}


// ======================================================
// 未返済一覧
// ======================================================

function listUnpaid_(sheet, talkId) {
  const values = getDataRows_(sheet);

  const unpaid = values.filter(row => {
    const status = String(row[COL_STATUS - 1] || '').trim();
    const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();

    return rowTalkId === talkId && status === '未返済';
  });

  if (unpaid.length === 0) {
    return 'このトークには未返済のデータはありません。';
  }

  const lines = unpaid
    .slice(0, 10)
    .map(row => {
      const id = row[COL_ID - 1];
      const person = row[COL_PERSON - 1];
      const amount = row[COL_AMOUNT - 1];
      const memo = row[COL_MEMO - 1];

      return `${id}\n${person} / ${amount}円 / ${memo || '内容なし'}`;
    });

  const suffix =
    unpaid.length > 10
      ? `\n\n他に${unpaid.length - 10}件あります。`
      : '';

  return [
    '未返済一覧:',
    '',
    lines.join('\n\n'),
    suffix
  ].join('\n');
}


// ======================================================
// 既存データを現在のグループへ移行
// ======================================================

function migrateExistingRowsToCurrentGroup_(sheet, userId, sourceType, talkId) {
  if (sourceType !== 'group') {
    return [
      'この操作はLINEグループ内から実行してください。',
      '',
      'Botを対象グループに追加したあと、',
      '「既存データをこのグループに移行」',
      'と送信してください。'
    ].join('\n');
  }

  if (!talkId) {
    return 'グループIDを取得できませんでした。';
  }

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(10000);

    const values = getDataRows_(sheet);
    const targetRows = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const id = String(row[COL_ID - 1] || '').trim();
      const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();

      if (id && !rowTalkId) {
        targetRows.push(i);
      }
    }

    if (targetRows.length === 0) {
      return '移行対象の既存データはありません。';
    }

    if (!canMigrateExistingRows_(values, userId)) {
      return [
        '既存データを移行する権限がありません。',
        '',
        '既存データの登録者本人から実行してください。'
      ].join('\n');
    }

    const lastRow = sheet.getLastRow();
    const talkValues =
      sheet
        .getRange(2, COL_TALK_TYPE, lastRow - 1, 2)
        .getValues();

    let count = 0;

    for (let i = 0; i < talkValues.length; i++) {
      const id = String(values[i][COL_ID - 1] || '').trim();
      const existingTalkId = String(talkValues[i][1] || '').trim();

      if (id && !existingTalkId) {
        talkValues[i][0] = 'group';
        talkValues[i][1] = talkId;
        count++;
      }
    }

    sheet
      .getRange(2, COL_TALK_TYPE, talkValues.length, 2)
      .setValues(talkValues);

    return [
      '既存データのグループ移行が完了しました。',
      '',
      `移行件数: ${count}件`,
      '',
      '次に各メンバーが自分の名前を登録してください。',
      '例: 名前登録 田中'
    ].join('\n');

  } catch (err) {
    console.error(err);
    return '既存データの移行中にエラーが発生しました。';

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // 何もしない
    }
  }
}


function canMigrateExistingRows_(values, userId) {
  const migrationAdminUserId =
    PropertiesService
      .getScriptProperties()
      .getProperty('MIGRATION_ADMIN_USER_ID');

  if (migrationAdminUserId) {
    return userId === migrationAdminUserId;
  }

  const registeredUserIds = new Set();

  values.forEach(row => {
    const id = String(row[COL_ID - 1] || '').trim();
    const registeredBy = String(row[COL_REGISTERED_BY - 1] || '').trim();
    const rowTalkId = String(row[COL_TALK_ID - 1] || '').trim();

    if (id && !rowTalkId && registeredBy) {
      registeredUserIds.add(registeredBy);
    }
  });

  if (registeredUserIds.size !== 1) {
    return false;
  }

  return Array.from(registeredUserIds)[0] === userId;
}


// ======================================================
// ヘルプ
// ======================================================

function helpText_() {
  return [
    '使い方:',
    '',
    '【貸したとき】',
    '貸した 相手 金額 内容',
    '例: 貸した 田中 3000 昼ごはん',
    '',
    '【返済されたとき】',
    '返した 相手 金額',
    '例: 返した 田中 3000',
    '',
    '【名前とLINE IDを紐付け】',
    '名前登録 自分の名前',
    '例: 名前登録 田中',
    '',
    '過去に別表記がある場合はその名前も登録:',
    '例: 名前登録 タナカ',
    '',
    '【登録名の確認】',
    '名前一覧',
    '',
    '【未返済一覧】',
    '未返済'
  ].join('\n');
}


// ======================================================
// シート / データ取得
// ======================================================

function getSheet_() {
  const ss = SpreadsheetApp.openById(getSpreadsheetId_());
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  return sheet;
}


function getDataRows_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  return sheet
    .getRange(2, 1, lastRow - 1, COL_REMINDER_COUNT)
    .getValues();
}


// ======================================================
// 貸付ID / 日時 / 正規化
// ======================================================

function createLoanId_() {
  const now = new Date();
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const hh = String(now.getHours()).padStart(2, '0');
  const mi = String(now.getMinutes()).padStart(2, '0');
  const ss = String(now.getSeconds()).padStart(2, '0');
  const rand = String(Math.floor(Math.random() * 10000)).padStart(4, '0');

  return `L${yyyy}${mm}${dd}${hh}${mi}${ss}${rand}`;
}


function nowString_() {
  return Utilities.formatDate(
    new Date(),
    'Asia/Tokyo',
    'yyyy/MM/dd HH:mm:ss'
  );
}


function normalizeText_(text) {
  return String(text)
    .replace(
      /[０-９]/g,
      s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)
    )
    .replace(/，/g, ',')
    .replace(/　/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function parseAmount_(amountText) {
  return Number(String(amountText).replace(/,/g, ''));
}



// ======================================================
// 未返済の自動督促通知
// ======================================================
//
// 仕様:
// - 貸付から60日以上経過し、未返済なら初回通知
// - 初回通知後も未返済なら7日ごとに再通知
// - 同じ借り手の通知対象が複数ある場合は1通にまとめる
// - 通知成功時のみ N列「最終督促通知日時」と
//   O列「督促回数」を更新する
//
// 必要な初期操作:
// 1. setupSheet() を1回実行
// 2. previewOverdueNotifications() で対象を確認
// 3. setupOverdueNotificationTrigger() を1回実行
//
// ======================================================

function getOverdueFirstDays_() {
  const value =
    PropertiesService
      .getScriptProperties()
      .getProperty('OVERDUE_FIRST_DAYS');

  const parsed = Number(value);

  return Number.isFinite(parsed) && parsed > 0
    ? Math.floor(parsed)
    : 60;
}


function getOverdueRepeatDays_() {
  const value =
    PropertiesService
      .getScriptProperties()
      .getProperty('OVERDUE_REPEAT_DAYS');

  const parsed = Number(value);

  return Number.isFinite(parsed) && parsed > 0
    ? Math.floor(parsed)
    : 7;
}


function getOverdueTriggerHour_() {
  const value =
    PropertiesService
      .getScriptProperties()
      .getProperty('OVERDUE_TRIGGER_HOUR');

  const parsed = Number(value);

  if (
    Number.isFinite(parsed) &&
    parsed >= 0 &&
    parsed <= 23
  ) {
    return Math.floor(parsed);
  }

  // デフォルトは朝9時台
  return 9;
}


// ------------------------------------------------------
// 毎日1回の時間主導トリガーを作成
// ------------------------------------------------------
//
// Apps Script上で1回だけ手動実行する。
// 既存の同名トリガーがあれば削除して作り直す。
// ------------------------------------------------------

function setupOverdueNotificationTrigger() {
  setupSheet();

  const handlerName = 'checkOverdueLoans';

  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      if (trigger.getHandlerFunction() === handlerName) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

  const hour = getOverdueTriggerHour_();

  ScriptApp
    .newTrigger(handlerName)
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();

  console.log(
    '未返済督促トリガーを作成しました。' +
    '毎日 ' + hour + '時台に checkOverdueLoans() を実行します。'
  );
}


// ------------------------------------------------------
// トリガー削除
// ------------------------------------------------------

function removeOverdueNotificationTrigger() {
  const handlerName = 'checkOverdueLoans';
  let count = 0;

  ScriptApp
    .getProjectTriggers()
    .forEach(trigger => {
      if (trigger.getHandlerFunction() === handlerName) {
        ScriptApp.deleteTrigger(trigger);
        count++;
      }
    });

  console.log(
    '未返済督促トリガーを ' + count + ' 件削除しました。'
  );
}


// ------------------------------------------------------
// 実際に通知を送信するメイン処理
// ------------------------------------------------------

function checkOverdueLoans() {
  setupSheet();

  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    const sheet = getSheet_();
    const rows = getDataRows_(sheet);
    const notificationGroups =
      collectOverdueNotificationGroups_(rows);

    const groupKeys = Object.keys(notificationGroups);

    if (groupKeys.length === 0) {
      console.log('今回の督促通知対象はありません。');
      return;
    }

    const now = new Date();
    const nowString = Utilities.formatDate(
      now,
      'Asia/Tokyo',
      'yyyy/MM/dd HH:mm:ss'
    );

    let sentGroupCount = 0;
    let updatedLoanCount = 0;

    groupKeys.forEach(key => {
      const group = notificationGroups[key];
      const message = buildOverdueMessage_(group);

      const sent = pushTextMessage_(
        group.talkId,
        message
      );

      if (!sent) {
        console.error(
          '督促通知の送信に失敗したため、通知履歴は更新しません: ' +
          key
        );
        return;
      }

      sentGroupCount++;

      group.items.forEach(item => {
        sheet
          .getRange(
            item.rowNumber,
            COL_LAST_REMINDER_AT
          )
          .setValue(nowString);

        sheet
          .getRange(
            item.rowNumber,
            COL_REMINDER_COUNT
          )
          .setValue(item.reminderCount + 1);

        updatedLoanCount++;
      });
    });

    console.log(
      '督促通知完了: ' +
      sentGroupCount +
      '通 / ' +
      updatedLoanCount +
      '件更新'
    );

  } catch (err) {
    console.error('checkOverdueLoans error');
    console.error(err);
    throw err;

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // 何もしない
    }
  }
}


// ------------------------------------------------------
// 送信せず、現在の対象だけログで確認
// ------------------------------------------------------
//
// 初めて自動通知を有効にする前に実行推奨。
// 既存の古い未返済データが一気に通知対象になるため、
// まずこの関数で内容を確認する。
// ------------------------------------------------------

function previewOverdueNotifications() {
  setupSheet();

  const sheet = getSheet_();
  const rows = getDataRows_(sheet);
  const notificationGroups =
    collectOverdueNotificationGroups_(rows);

  const keys = Object.keys(notificationGroups);

  if (keys.length === 0) {
    console.log('現在、督促通知対象はありません。');
    return;
  }

  console.log(
    '督促通知対象グループ数: ' + keys.length
  );

  keys.forEach((key, index) => {
    const group = notificationGroups[key];

    console.log(
      '===== 通知プレビュー ' +
      (index + 1) +
      ' ====='
    );

    console.log(buildOverdueMessage_(group));
  });
}


// ------------------------------------------------------
// 通知対象を抽出し、
// talkId + 借り手LINE ID（未登録なら名前）単位にまとめる
// ------------------------------------------------------

function collectOverdueNotificationGroups_(rows) {
  const firstDays = getOverdueFirstDays_();
  const repeatDays = getOverdueRepeatDays_();
  const now = new Date();
  const todayIndex = tokyoDayIndex_(now);

  const groups = {};

  rows.forEach((row, index) => {
    const status =
      String(row[COL_STATUS - 1] || '').trim();

    if (status !== '未返済') {
      return;
    }

    const talkType =
      String(row[COL_TALK_TYPE - 1] || '').trim();

    const talkId =
      String(row[COL_TALK_ID - 1] || '').trim();

    // 今回の自動通知はLINEグループ宛てのみ
    if (talkType !== 'group' || !talkId) {
      return;
    }

    const lentAt =
      parseSheetDateValue_(row[COL_LENT_AT - 1]);

    if (!lentAt) {
      console.error(
        '貸付日時を解析できません: row=' +
        (index + 2)
      );
      return;
    }

    const elapsedDays =
      todayIndex - tokyoDayIndex_(lentAt);

    if (elapsedDays < firstDays) {
      return;
    }

    const reminderCount =
      Number(row[COL_REMINDER_COUNT - 1]) || 0;

    const lastReminderAt =
      parseSheetDateValue_(
        row[COL_LAST_REMINDER_AT - 1]
      );

    let shouldNotify = false;

    // 初回
    if (reminderCount <= 0 || !lastReminderAt) {
      shouldNotify = true;

    } else {
      // 再通知
      const daysSinceLastReminder =
        todayIndex - tokyoDayIndex_(lastReminderAt);

      shouldNotify =
        daysSinceLastReminder >= repeatDays;
    }

    if (!shouldNotify) {
      return;
    }

    const rawBorrowerName =
      String(row[COL_PERSON - 1] || '').trim() ||
      '未設定';

    let borrowerUserId =
      String(
        row[COL_BORROWER_USER_ID - 1] || ''
      ).trim();

    // 古いデータでM列が空なら名前から解決
    if (
      !borrowerUserId &&
      typeof resolveBorrowerUserId_ === 'function'
    ) {
      borrowerUserId =
        resolveBorrowerUserId_(
          talkId,
          rawBorrowerName
        );

      if (borrowerUserId) {
        getSheet_()
          .getRange(
            index + 2,
            COL_BORROWER_USER_ID
          )
          .setValue(borrowerUserId);
      }
    }

    const borrowerKey =
      borrowerUserId ||
      ('name:' + normalizeAliasKey_(rawBorrowerName));

    const key =
      talkId + '::' + borrowerKey;

    if (!groups[key]) {
      groups[key] = {
        talkId: talkId,
        talkType: talkType,
        borrowerUserId: borrowerUserId,
        borrowerName: borrowerUserId
          ? getMemberDisplayNameForBot_(
              talkType,
              talkId,
              borrowerUserId
            )
          : rawBorrowerName,
        items: []
      };
    }

    groups[key].items.push({
      rowNumber: index + 2,
      loanId:
        String(row[COL_ID - 1] || '').trim(),
      amount:
        Number(row[COL_AMOUNT - 1]) || 0,
      memo:
        String(row[COL_MEMO - 1] || '').trim(),
      lentAt: lentAt,
      elapsedDays: elapsedDays,
      reminderCount: reminderCount
    });
  });

  return groups;
}


// ------------------------------------------------------
// 通知文を生成
// ------------------------------------------------------

function buildOverdueMessage_(group) {
  const firstDays = getOverdueFirstDays_();
  const repeatDays = getOverdueRepeatDays_();

  const sortedItems =
    group.items
      .slice()
      .sort((a, b) => {
        if (b.elapsedDays !== a.elapsedDays) {
          return b.elapsedDays - a.elapsedDays;
        }
        return b.amount - a.amount;
      });

  const total =
    sortedItems.reduce(
      (sum, item) => sum + item.amount,
      0
    );

  const maxShown = 10;
  const shown = sortedItems.slice(0, maxShown);

  const lines = [
    '⚠️ 未返済のお知らせ',
    '',
    `${group.borrowerName}さんに、` +
      `${firstDays}日以上経過した未返済があります。`,
    ''
  ];

  shown.forEach(item => {
    const lentDate =
      Utilities.formatDate(
        item.lentAt,
        'Asia/Tokyo',
        'yyyy/MM/dd'
      );

    lines.push(
      `・${formatYen_(item.amount)}` +
      (item.memo ? ` / ${item.memo}` : ''),
      `  貸付日: ${lentDate}` +
      `（${item.elapsedDays}日経過）`
    );
  });

  if (sortedItems.length > maxShown) {
    lines.push(
      `ほか${sortedItems.length - maxShown}件`
    );
  }

  lines.push(
    '',
    `合計: ${formatYen_(total)}`,
    '',
    '返済済みの場合は、LINEまたはWebアプリから更新してください。',
    `未返済のままの場合は${repeatDays}日後に再通知します。`
  );

  return lines.join('\n');
}


// ------------------------------------------------------
// LINE Push API
// ------------------------------------------------------

function pushTextMessage_(to, text) {
  const token = getChannelAccessToken_();

  if (!token) {
    console.error(
      'CHANNEL_ACCESS_TOKEN が設定されていません。'
    );
    return false;
  }

  const target = String(to || '').trim();

  if (!target) {
    console.error('Push通知先のtalkIdが空です。');
    return false;
  }

  const url =
    'https://api.line.me/v2/bot/message/push';

  const payload = {
    to: target,
    messages: [
      {
        type: 'text',
        text: String(text || '').slice(0, 4900)
      }
    ]
  };

  try {
    const response = UrlFetchApp.fetch(
      url,
      {
        method: 'post',
        contentType: 'application/json',
        headers: {
          Authorization: `Bearer ${token}`
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    const code = response.getResponseCode();

    if (code >= 200 && code < 300) {
      return true;
    }

    console.error(
      'LINE Push API error: ' +
      code +
      ' ' +
      response.getContentText()
    );

    return false;

  } catch (err) {
    console.error('LINE Push API exception');
    console.error(err);
    return false;
  }
}


// ------------------------------------------------------
// 日付処理
// ------------------------------------------------------

function parseSheetDateValue_(value) {
  if (!value) {
    return null;
  }

  if (
    Object.prototype.toString.call(value) === '[object Date]' &&
    !isNaN(value.getTime())
  ) {
    return value;
  }

  const text =
    String(value).trim();

  // yyyy/MM/dd H:mm:ss
  // yyyy/MM/dd HH:mm:ss
  const match =
    text.match(
      /^(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/
    );

  if (match) {
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const hour = Number(match[4]);
    const minute = Number(match[5]);
    const second = Number(match[6] || 0);

    // 日本時間としてDate化
    return new Date(
      Date.UTC(
        year,
        month - 1,
        day,
        hour - 9,
        minute,
        second
      )
    );
  }

  const parsed = new Date(text);

  if (isNaN(parsed.getTime())) {
    return null;
  }

  return parsed;
}


function tokyoDayIndex_(date) {
  const parts =
    Utilities
      .formatDate(
        date,
        'Asia/Tokyo',
        'yyyy,M,d'
      )
      .split(',')
      .map(Number);

  return Math.floor(
    Date.UTC(
      parts[0],
      parts[1] - 1,
      parts[2]
    ) /
    (24 * 60 * 60 * 1000)
  );
}


function formatYen_(amount) {
  const value =
    Math.round(Number(amount) || 0);

  return (
    value
      .toLocaleString('ja-JP') +
    '円'
  );
}



// ======================================================
// LINEプロフィール取得（Bot表示用）
// ======================================================

function getMemberDisplayNameForBot_(sourceType, talkId, userId) {
  if (!userId) return '不明';

  const token = getChannelAccessToken_();
  if (!token) return `メンバー ${String(userId).slice(-4)}`;

  let url = '';

  if (sourceType === 'group') {
    url =
      'https://api.line.me/v2/bot/group/' +
      encodeURIComponent(talkId) +
      '/member/' +
      encodeURIComponent(userId);
  } else if (sourceType === 'room') {
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
        Authorization: `Bearer ${token}`
      },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      return data.displayName || `メンバー ${String(userId).slice(-4)}`;
    }
  } catch (err) {
    console.error(err);
  }

  return `メンバー ${String(userId).slice(-4)}`;
}


// ======================================================
// LINE返信
// ======================================================

function getChannelAccessToken_() {
  return PropertiesService
    .getScriptProperties()
    .getProperty('CHANNEL_ACCESS_TOKEN') || '';
}


function replyToLine_(replyToken, text) {
  const channelAccessToken = getChannelAccessToken_();

  if (!channelAccessToken) {
    console.error('CHANNEL_ACCESS_TOKEN が設定されていません。');
    return;
  }

  const url = 'https://api.line.me/v2/bot/message/reply';

  const payload = {
    replyToken: replyToken,
    messages: [
      {
        type: 'text',
        text: text
      }
    ]
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${channelAccessToken}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();

  if (code < 200 || code >= 300) {
    console.error(
      'LINE Reply API error: ' +
      code +
      ' ' +
      response.getContentText()
    );
  }
}

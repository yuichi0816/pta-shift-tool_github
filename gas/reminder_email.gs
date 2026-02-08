/**
 * PTA旗振りシフト リマインドメール送信スクリプト
 * 
 * 使い方:
 * 1. このスクリプトをGoogleスプレッドシートのApps Scriptにコピー
 * 2. 「設定」シートを作成し、必要な情報を入力
 * 3. sendReminders関数を毎日実行するトリガーを設定
 */

// ╔════════════════════════════════════════════════════════════════╗
// ║                    ★ 設定エリア ★                              ║
// ║  ここを環境に合わせて変更してください                            ║
// ╚════════════════════════════════════════════════════════════════╝

const CONFIG = {
  
  // ─────────────────────────────────────────────────────────────────
  // シート名の設定
  // ─────────────────────────────────────────────────────────────────
  SHIFT_SHEET_NAME: 'シフト結果',       // シフト結果が入っているシート名
  SURVEY_SHEET_NAME: 'アンケート回答',   // アンケート回答が入っているシート名
  LOG_SHEET_NAME: '送信ログ',            // 送信ログを記録するシート名（自動作成）
  
  // ─────────────────────────────────────────────────────────────────
  // シフト結果シートの列設定（横持ち形式）
  // 
  // 例: | 日付 | 曜日 | 地点A | 地点B | 地点C | ...
  //     | 2/5  | 月   | 田中  | 佐藤  | 山田  | ...
  // ─────────────────────────────────────────────────────────────────
  SHIFT_DATE_COL: 'A',              // 日付が入っている列
  SHIFT_LOCATION_START_COL: 'C',    // 地点の開始列
  // ※地点の終了列は自動検出（ヘッダー行の空でない最後の列まで）
  
  // ─────────────────────────────────────────────────────────────────
  // アンケート回答シートの列設定
  // ─────────────────────────────────────────────────────────────────
  SURVEY_EMAIL_COL: 'B',            // メールアドレス列
  SURVEY_NAME_COL: 'N',             // 氏名列
  
  // ─────────────────────────────────────────────────────────────────
  // 氏名の照合設定
  // 
  // シフト結果の氏名に「_2年」や「_4年(2)」のようなサフィックスが付いている場合、
  // この正規表現で除去してからアンケート回答と照合します。
  // 例: 「佐藤太郎_2年」「佐藤太郎_4年(2)」→「佐藤太郎」として照合
  // ─────────────────────────────────────────────────────────────────
  NAME_SUFFIX_PATTERN: /_\d+年(\(\d+\))?$/,  // _1年, _2年, _4年(2) などにマッチ
  
  // ─────────────────────────────────────────────────────────────────
  // リマインド設定
  // ─────────────────────────────────────────────────────────────────
  DAYS_BEFORE: 1,                   // 何日前にリマインドするか（1=前日）
  SEND_HOUR: 20,                    // 何時に送信するか（0-23、8時）
  
  // ─────────────────────────────────────────────────────────────────
  // メール設定
  // ─────────────────────────────────────────────────────────────────
  SUBJECT: 'これはテストです【リマインド】明日は旗振り当番です',
  LABEL_NAME: '旗振りリマインド',   // Gmailに付けるラベル名
  
  // ─────────────────────────────────────────────────────────────────
  // 添付ファイル設定
  // スプレッドシートと同じフォルダにあるPDFを添付
  // ファイル名の末尾が一致するファイルを検索
  // 添付しない場合は '' （空文字列）に設定
  // ─────────────────────────────────────────────────────────────────
  ATTACHMENT_FILENAME: 'R8新入生向け_旗振り地点について.pdf',
};

// ─────────────────────────────────────────────────────────────────
// 列名変換ヘルパー（A→0, B→1, C→2, ... AA→26, AB→27, ...）
// ─────────────────────────────────────────────────────────────────
function colToIndex(colLetter) {
  const col = colLetter.toUpperCase();
  let index = 0;
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64);
  }
  return index - 1; // 0始まりのインデックスに変換
}

// ==================== メイン処理 ====================

/**
 * リマインドメールを送信するメイン関数
 * トリガーからこの関数を呼び出す
 */
function sendReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 各シートを取得
  const shiftSheet = ss.getSheetByName(CONFIG.SHIFT_SHEET_NAME);
  const surveySheet = ss.getSheetByName(CONFIG.SURVEY_SHEET_NAME);
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  
  // シートの存在チェック
  if (!shiftSheet) {
    Logger.log('エラー: シフト結果シートが見つかりません');
    return;
  }
  if (!surveySheet) {
    Logger.log('エラー: アンケート回答シートが見つかりません');
    return;
  }
  
  // 送信ログシートがなければ作成
  if (!logSheet) {
    logSheet = createLogSheet(ss);
  }
  
  // 明日の日付を取得
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + CONFIG.DAYS_BEFORE);
  const tomorrowStr = formatDate(tomorrow);
  
  Logger.log(`対象日付: ${tomorrowStr}`);
  
  // シフトデータから明日の担当者を抽出
  const assignments = getTomorrowAssignments(shiftSheet, tomorrow);
  Logger.log(`明日の担当者数: ${assignments.length}`);
  
  if (assignments.length === 0) {
    Logger.log('明日の担当者はいません');
    return;
  }
  
  // メールアドレスマップを作成
  const emailMap = createEmailMap(surveySheet);
  
  // ラベルを取得または作成
  const label = getOrCreateLabel(CONFIG.LABEL_NAME);
  
  // 各担当者にメール送信
  let sentCount = 0;
  for (const assignment of assignments) {
    // シフト結果の氏名からサフィックスを除去して照合
    const nameForLookup = stripNameSuffix(assignment.name);
    const email = emailMap[normalizeString(nameForLookup)];
    
    if (!email) {
      Logger.log(`警告: ${assignment.name} のメールアドレスが見つかりません`);
      logToSheet(logSheet, tomorrowStr, assignment.name, assignment.location, '', '送信失敗: メールアドレス不明');
      continue;
    }
    
    // 二重送信チェック
    if (isAlreadySent(logSheet, tomorrowStr, email)) {
      Logger.log(`スキップ: ${assignment.name} には既に送信済み`);
      continue;
    }
    
    // メール送信
    try {
      sendReminderEmail(email, assignment, label);
      logToSheet(logSheet, tomorrowStr, assignment.name, assignment.location, email, '送信成功');
      sentCount++;
      Logger.log(`送信完了: ${assignment.name} (${email})`);
    } catch (e) {
      Logger.log(`エラー: ${assignment.name} への送信失敗 - ${e.message}`);
      logToSheet(logSheet, tomorrowStr, assignment.name, assignment.location, email, `送信失敗: ${e.message}`);
    }
  }
  
  Logger.log(`送信完了: ${sentCount}件`);
}

// ==================== ヘルパー関数 ====================

/**
 * 明日のシフト担当者を取得（横持ち形式対応）
 * シート構造: A列=日付, B列=曜日, C列以降=地点（ヘッダー行に地点名）
 * 地点の終了列はヘッダー行から自動検出
 */
function getTomorrowAssignments(sheet, targetDate) {
  const data = sheet.getDataRange().getValues();
  const assignments = [];
  
  // ヘッダー行から地点名を取得（空でない列まで）
  const headerRow = data[0];
  const locationNames = [];
  const startCol = colToIndex(CONFIG.SHIFT_LOCATION_START_COL);
  
  // ヘッダー行を走査して地点を収集
  for (let col = startCol; col < headerRow.length; col++) {
    const cellValue = headerRow[col];
    if (cellValue && String(cellValue).trim()) {
      locationNames.push(String(cellValue).trim());
    } else {
      // 空セルに到達したら終了
      break;
    }
  }
  
  Logger.log(`検出した地点数: ${locationNames.length} (${locationNames.join(', ')})`);
  
  // 各行をチェック
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateValue = row[colToIndex(CONFIG.SHIFT_DATE_COL)];
    
    if (!dateValue) continue;
    
    // 日付を比較
    const shiftDate = new Date(dateValue);
    if (isSameDate(shiftDate, targetDate)) {
      // この日付の行の各地点をチェック
      for (let j = 0; j < locationNames.length; j++) {
        const col = startCol + j;
        const name = row[col];
        if (name && String(name).trim()) {
          // カンマ区切りで複数名いる場合は分割
          const names = String(name).split(/[,、]/);
          for (const n of names) {
            if (n.trim()) {
              assignments.push({
                date: formatDate(targetDate),
                location: locationNames[j],
                name: n.trim()
              });
            }
          }
        }
      }
      break; // 対象日付が見つかったらループ終了
    }
  }
  
  return assignments;
}

/**
 * メールアドレスマップを作成（氏名 -> メールアドレス）
 */
function createEmailMap(sheet) {
  const data = sheet.getDataRange().getValues();
  const map = {};
  
  for (let i = 1; i < data.length; i++) { // ヘッダー行をスキップ
    const row = data[i];
    const email = row[colToIndex(CONFIG.SURVEY_EMAIL_COL)];
    const name = row[colToIndex(CONFIG.SURVEY_NAME_COL)];
    
    if (email && name) {
      map[normalizeString(name)] = email;
    }
  }
  
  return map;
}

/**
 * リマインドメールを送信
 */
function sendReminderEmail(email, assignment, label) {
  const body = `
${assignment.name} 様
動作確認メールです。4月以降、こんなメールが届くことになります。

明日は旗振り当番の日です。

━━━━━━━━━━━━━━━━━━━━━━
📅 日付: ${assignment.date}
📍 場所: ${assignment.location}
━━━━━━━━━━━━━━━━━━━━━━

お忙しいところ恐れ入りますが、
ご対応のほどよろしくお願いいたします。

※このメールは自動送信されています。
※ご都合が悪くなった場合は、PTA担当者までご連絡ください。

--
PTA旗振りシフト管理システム
`;

  // 添付ファイルを取得
  const attachment = getAttachmentFile();
  
  // メール送信オプション
  const options = {};
  if (attachment) {
    options.attachments = [attachment];
  }
  
  // メール送信
  GmailApp.sendEmail(email, CONFIG.SUBJECT, body, options);
  
  // 送信済みメールにラベルを付けてアーカイブ
  Utilities.sleep(2000); // 送信が完了するまで少し待機
  
  const threads = GmailApp.search(`to:${email} subject:"${CONFIG.SUBJECT}" newer_than:1d`, 0, 1);
  if (threads.length > 0) {
    const thread = threads[0];
    thread.addLabel(label);
    thread.moveToArchive();
  }
}

/**
 * 添付ファイルを取得（スプレッドシートと同じフォルダから）
 * ファイル名の末尾が一致するファイルを検索
 */
function getAttachmentFile() {
  // 添付ファイル名が設定されていない場合はスキップ
  if (!CONFIG.ATTACHMENT_FILENAME) {
    return null;
  }
  
  try {
    // スプレッドシートの親フォルダを取得
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssFile = DriveApp.getFileById(ss.getId());
    const parents = ssFile.getParents();
    
    if (!parents.hasNext()) {
      Logger.log('警告: スプレッドシートの親フォルダが見つかりません');
      return null;
    }
    
    const folder = parents.next();
    
    // フォルダ内の全ファイルを走査し、末尾が一致するファイルを検索
    const allFiles = folder.getFiles();
    let foundFile = null;
    
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      const fileName = file.getName();
      
      // ファイル名の末尾が設定値と一致するかチェック
      if (fileName.endsWith(CONFIG.ATTACHMENT_FILENAME)) {
        foundFile = file;
        break;
      }
    }
    
    if (!foundFile) {
      Logger.log(`警告: "*${CONFIG.ATTACHMENT_FILENAME}" に一致する添付ファイルが見つかりません`);
      return null;
    }
    
    Logger.log(`添付ファイル: ${foundFile.getName()}`);
    return foundFile.getBlob();
    
  } catch (e) {
    Logger.log(`添付ファイル取得エラー: ${e.message}`);
    return null;
  }
}

/**
 * ラベルを取得または作成
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) {
    label = GmailApp.createLabel(labelName);
  }
  return label;
}

/**
 * 送信ログシートを作成
 */
function createLogSheet(ss) {
  const sheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  sheet.appendRow(['送信日時', '担当日', '担当者名', '場所', 'メールアドレス', 'ステータス']);
  
  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, 6);
  headerRange.setBackground('#4a86e8');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  return sheet;
}

/**
 * 送信ログに記録
 */
function logToSheet(sheet, targetDate, name, location, email, status) {
  sheet.appendRow([
    new Date(),
    targetDate,
    name,
    location,
    email,
    status
  ]);
}

/**
 * 既に送信済みかチェック
 */
function isAlreadySent(sheet, targetDate, email) {
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[1] === targetDate && row[4] === email && row[5] === '送信成功') {
      return true;
    }
  }
  
  return false;
}

/**
 * 日付をYYYY/MM/DD形式にフォーマット
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}/${month}/${day}`;
}

/**
 * 同じ日付かどうかをチェック
 */
function isSameDate(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

/**
 * 氏名からサフィックス（_2年など）を除去
 * CONFIG.NAME_SUFFIX_PATTERN に設定されたパターンを使用
 */
function stripNameSuffix(name) {
  if (!name) return '';
  return String(name).replace(CONFIG.NAME_SUFFIX_PATTERN, '');
}

/**
 * 文字列を正規化（空白削除、全角半角統一）
 */
function normalizeString(str) {
  if (!str) return '';
  return String(str)
    .trim()
    .replace(/\s+/g, '')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
}

// ==================== セットアップ関数 ====================

/**
 * 毎日のトリガーを設定
 */
function createDailyTrigger() {
  // 既存のトリガーを削除
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendReminders') {
      ScriptApp.deleteTrigger(trigger);
    }
  }
  
  // 新しいトリガーを作成
  ScriptApp.newTrigger('sendReminders')
    .timeBased()
    .atHour(CONFIG.SEND_HOUR)
    .everyDays(1)
    .create();
  
  Logger.log(`トリガーを設定しました: 毎日 ${CONFIG.SEND_HOUR}時に実行`);
}

/**
 * 🗑️ 毎日のトリガーを削除（自動送信を停止）
 */
function deleteDailyTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'sendReminders') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  }
  
  if (deletedCount > 0) {
    Logger.log(`${deletedCount}件のトリガーを削除しました。自動送信は停止されました。`);
  } else {
    Logger.log('削除するトリガーはありませんでした。');
  }
}

/**
 * テスト送信（今日の担当者にテストメール送信）
 */
function testSendReminders() {
  // 設定を一時的に変更してテスト
  CONFIG.DAYS_BEFORE = 0; // 今日の担当者を対象
  sendReminders();
}

// ==================== シミュレーション ====================

/**
 * 🔍 シミュレーション実行（メール送信なし）
 * 
 * 実際にメールを送信せずに、以下を確認できます：
 * - シフトデータの読み取り
 * - メールアドレスの照合
 * - 送信対象者の一覧
 * 
 * 結果は「表示」→「ログ」で確認してください
 */
function simulateReminders() {
  simulateForDate(CONFIG.DAYS_BEFORE);
}

/**
 * 🔍 特定日付でシミュレーション
 * @param {number} daysFromToday - 今日から何日後（0=今日、1=明日、-1=昨日）
 */
function simulateForDate(daysFromToday) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log('════════════════════════════════════════════════════════');
  Logger.log('📧 リマインドメール シミュレーション（メール送信なし）');
  Logger.log('════════════════════════════════════════════════════════');
  
  // シートを取得
  const shiftSheet = ss.getSheetByName(CONFIG.SHIFT_SHEET_NAME);
  const surveySheet = ss.getSheetByName(CONFIG.SURVEY_SHEET_NAME);
  
  if (!shiftSheet) {
    Logger.log(`❌ エラー: "${CONFIG.SHIFT_SHEET_NAME}" シートが見つかりません`);
    return;
  }
  if (!surveySheet) {
    Logger.log(`❌ エラー: "${CONFIG.SURVEY_SHEET_NAME}" シートが見つかりません`);
    return;
  }
  
  Logger.log(`✅ シフト結果シート: ${CONFIG.SHIFT_SHEET_NAME}`);
  Logger.log(`✅ アンケート回答シート: ${CONFIG.SURVEY_SHEET_NAME}`);
  
  // 対象日付を計算
  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + daysFromToday);
  const targetDateStr = formatDate(targetDate);
  
  Logger.log('');
  Logger.log(`📅 対象日付: ${targetDateStr}（今日から${daysFromToday}日後）`);
  Logger.log('');
  
  // シフトデータから担当者を抽出
  const assignments = getTomorrowAssignments(shiftSheet, targetDate);
  
  if (assignments.length === 0) {
    Logger.log('⚠️ この日の担当者は見つかりませんでした');
    return;
  }
  
  Logger.log(`👥 担当者数: ${assignments.length}人`);
  Logger.log('');
  
  // メールアドレスマップを作成
  const emailMap = createEmailMap(surveySheet);
  Logger.log(`📋 アンケート回答から取得したメールアドレス数: ${Object.keys(emailMap).length}件`);
  Logger.log('');
  
  // 各担当者の照合結果を表示
  Logger.log('────────────────────────────────────────────────────────');
  Logger.log('照合結果:');
  Logger.log('────────────────────────────────────────────────────────');
  
  let matchCount = 0;
  let noMatchCount = 0;
  
  for (const assignment of assignments) {
    const originalName = assignment.name;
    const nameForLookup = stripNameSuffix(originalName);
    const normalizedName = normalizeString(nameForLookup);
    const email = emailMap[normalizedName];
    
    if (email) {
      Logger.log(`✅ ${originalName}`);
      Logger.log(`   → 照合名: "${nameForLookup}" → メール: ${email}`);
      Logger.log(`   → 場所: ${assignment.location}`);
      matchCount++;
    } else {
      Logger.log(`❌ ${originalName}`);
      Logger.log(`   → 照合名: "${nameForLookup}" → メールアドレス不明`);
      Logger.log(`   → 場所: ${assignment.location}`);
      noMatchCount++;
    }
    Logger.log('');
  }
  
  // サマリー
  Logger.log('════════════════════════════════════════════════════════');
  Logger.log('📊 シミュレーション結果サマリー');
  Logger.log('════════════════════════════════════════════════════════');
  Logger.log(`対象日付: ${targetDateStr}`);
  Logger.log(`担当者数: ${assignments.length}人`);
  Logger.log(`  ✅ メール送信可能: ${matchCount}人`);
  Logger.log(`  ❌ メールアドレス不明: ${noMatchCount}人`);
  Logger.log('');
  Logger.log('※ このシミュレーションではメールは送信されていません');
  Logger.log('※ 結果を確認後、問題なければ createDailyTrigger() でトリガーを設定してください');
}

/**
 * 🔍 明日のシミュレーション
 */
function simulateTomorrow() {
  simulateForDate(1);
}

/**
 * 🔍 今日のシミュレーション
 */
function simulateToday() {
  simulateForDate(0);
}

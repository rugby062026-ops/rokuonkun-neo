/**
 * @license
 * このソフトウェアは、MITライセンスのもとで公開されています。
 * This software is released under the MIT License.
 *
 * Copyright (c) 2024 Masaaki Maeta
 *
 * 以下に定める条件に従い、本ソフトウェアおよび関連文書のファイル（以下「本ソフトウェア」）の複製を取得するすべての人に対し、本ソフトウェアを無制限に扱うことを無償で許可します。これには、本ソフトウェアの複製を使用、複写、変更、結合、掲載、頒布、サブライセンス、および/または販売する権利、および本ソフトウェアを提供する相手に同じことを許可する権利も無制限に含まれます。
 *
 * 上記の著作権表示および本許諾表示を、本ソフトウェアのすべての複製または重要な部分に記載するものとします。
 *
 * 本ソフトウェアは「現状のまま」で、明示であるか暗黙であるかを問わず、何らの保証もなく提供されます。ここでいう保証とは、商品性、特定の目的への適合性、および権利非侵害についての保証も含みますが、それに限定されるものではありません。 作者または著作権者は、契約行為、不法行為、またはそれ以外であろうと、本ソフトウェアに起因または関連し、あるいは本ソフトウェアの使用またはその他の扱いによって生じる一切の請求、損害、その他の義務について何らの責任も負わないものとします。
 *
 * --- (English Original) ---
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

// Code.gs - Web API として動作するサーバーサイドコード

const PARENT_FOLDER_NAME = "録音くん保存フォルダ";
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SUBFOLDER_SHEET_NAME = "シート1";
const SUBFOLDER_CELL = "B1";
const HISTORY_SHEET_NAME = "履歴";
const HISTORY_HEADERS = ["ファイル名", "保存日時", "フォルダパス", "ファイルリンク"];

/**
 * GETリクエストのハンドラ（ヘルスチェック用）
 */
function doGet(e) {
  const output = ContentService.createTextOutput(
    JSON.stringify({ status: "ok", message: "録音アプリ API は稼働中です。" })
  );
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * POSTリクエストのハンドラ（フロントエンドからのAPI呼び出し）
 */
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;

    let result;
    switch (action) {
      case "saveAudioFile":
        result = saveAudioFile(requestData.audioDataUrl, requestData.baseFileName);
        break;
      default:
        result = { success: false, message: "不明なアクションです: " + action };
    }

    const output = ContentService.createTextOutput(JSON.stringify(result));
    output.setMimeType(ContentService.MimeType.JSON);
    return output;

  } catch (error) {
    Logger.log("doPost エラー: " + error.toString());
    const output = ContentService.createTextOutput(
      JSON.stringify({ success: false, message: "サーバーエラー: " + error.message })
    );
    output.setMimeType(ContentService.MimeType.JSON);
    return output;
  }
}

function getOrCreateFolderIdByName(folderName, parentFolder = DriveApp.getRootFolder()) {
  if (!folderName || typeof folderName !== 'string' || folderName.trim() === '') {
    throw new Error("有効なフォルダ名が指定されていません。");
  }
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next().getId();
  } else {
    try {
      const newFolder = parentFolder.createFolder(folderName);
      Logger.log(`フォルダ "${folderName}" (ID: ${newFolder.getId()}) を親フォルダ "${parentFolder.getName()}" 内に作成しました。`);
      return newFolder.getId();
    } catch (error) {
      Logger.log(`フォルダ "${folderName}" の作成に失敗しました: ${error.toString()}`);
      throw new Error(`フォルダ "${folderName}" の作成に失敗しました。ドライブの権限などを確認してください。`);
    }
  }
}

function getSubFolderNameFromSheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SUBFOLDER_SHEET_NAME);
    if (!sheet) {
      Logger.log(`シート "${SUBFOLDER_SHEET_NAME}" が見つかりません。`);
      return null;
    }
    const subFolderName = sheet.getRange(SUBFOLDER_CELL).getValue().toString().trim();
    if (!subFolderName) {
      Logger.log(`セル "${SUBFOLDER_CELL}" にサブフォルダ名が入力されていません。`);
      return null;
    }
    return subFolderName.replace(/[\\\/:*?"<>|]/g, '_');
  } catch (e) {
    Logger.log(`スプレッドシートからのサブフォルダ名取得エラー: ${e.toString()}`);
    return null;
  }
}


/**
 * 履歴シートが存在しない場合に作成し、ヘッダーを書き込みます。
 */
function createHistorySheetIfNotExists() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.appendRow(HISTORY_HEADERS);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setFontWeight("bold");
    sheet.setColumnWidth(1, 250);
    sheet.setColumnWidth(2, 150);
    sheet.setColumnWidth(3, 250);
    sheet.setColumnWidth(4, 300);
    Logger.log(`履歴シート "${HISTORY_SHEET_NAME}" を作成しました。`);
  }
  return sheet;
}

/**
 * 履歴シートに録音情報を追記します。
 */
function addRecordToHistorySheet(fileName, folderPathText, folderUrl, fileUrl) {
  try {
    const sheet = createHistorySheetIfNotExists();
    const timestamp = new Date();
    const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "yyyy/MM/dd HH:mm:ss");

    const folderLinkFormula = `=HYPERLINK("${folderUrl}","${folderPathText}")`;
    const fileLinkFormula = `=HYPERLINK("${fileUrl}","${fileName}")`;

    sheet.appendRow([fileName, formattedTimestamp, folderLinkFormula, fileLinkFormula]);

    Logger.log(`履歴を記録しました: ${fileName}, ${formattedTimestamp}, フォルダ: ${folderPathText} (${folderUrl}), ファイル: ${fileUrl}`);
  } catch (e) {
    Logger.log(`履歴シートへの記録エラー: ${e.toString()}`);
  }
}


/**
 * Base64エンコードされた音声データ (Data URL形式) をデコードし、
 * 指定されたファイル名でGoogleドライブの指定フォルダにMP3ファイルとして保存し、履歴を記録します。
 */
function saveAudioFile(audioDataUrl, baseFileName) {
  try {
    if (!audioDataUrl || typeof audioDataUrl !== 'string') {
      throw new Error("音声データ (Data URL) が無効です。");
    }
    if (!baseFileName || typeof baseFileName !== 'string' || baseFileName.trim() === '') {
      throw new Error("ファイル名が無効です。");
    }

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);

    const subFolderNameRaw = getSubFolderNameFromSheet();
    let targetFolder;
    let folderPathText;
    let targetFolderUrl;

    if (subFolderNameRaw) {
      const subFolderId = getOrCreateFolderIdByName(subFolderNameRaw, parentFolder);
      targetFolder = DriveApp.getFolderById(subFolderId);
      folderPathText = `${parentFolder.getName()} > ${targetFolder.getName()}`;
      targetFolderUrl = targetFolder.getUrl();
    } else {
      targetFolder = parentFolder;
      folderPathText = parentFolder.getName();
      targetFolderUrl = parentFolder.getUrl();
      Logger.log(`サブフォルダ名が取得できなかったため、親フォルダ "${parentFolder.getName()}" に保存します。`);
    }

    const parts = audioDataUrl.match(/^data:(.+?);base64,(.+)$/);
    if (!parts || parts.length !== 3) {
      throw new Error("無効なData URL形式です。Base64エンコードされたデータを期待します。");
    }
    const mimeType = parts[1];
    const base64Data = parts[2];

    if (mimeType.indexOf('audio') === -1) {
        Logger.log(`警告: 期待されるMIMEタイプは audio/* ですが、${mimeType} を受信しました。`);
    }

    const finalFileName = baseFileName.toLowerCase().endsWith('.mp3')
                          ? baseFileName
                          : `${baseFileName}.mp3`;

    const decodedData = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decodedData, mimeType, finalFileName);

    const file = targetFolder.createFile(blob);
    const fileUrl = file.getUrl();

    addRecordToHistorySheet(finalFileName, folderPathText, targetFolderUrl, fileUrl);

    Logger.log(`ファイル "${finalFileName}" (ID: ${file.getId()}) をフォルダ "${folderPathText}" に保存しました。URL: ${fileUrl}`);
    return {
      success: true,
      message: `ファイル "${finalFileName}" をドライブのフォルダ「${folderPathText}」に保存し、履歴を記録しました。`,
      fileId: file.getId(),
      fileName: finalFileName,
      fileUrl: fileUrl
    };

  } catch (error) {
    Logger.log(`saveAudioFileでエラーが発生しました: ${error.toString()}\nStack: ${error.stack}`);
    let userFriendlyMessage = "ファイルの保存中にサーバー側でエラーが発生しました。";
    if (error.message.includes("Data URL")) {
        userFriendlyMessage = "受信した音声データの形式に問題があるようです。";
    } else if (error.message.includes("フォルダ")) {
        userFriendlyMessage = "保存先フォルダの準備中に問題が発生しました。";
    }
    return {
      success: false,
      message: `${userFriendlyMessage} (詳細は管理者に確認してください)`
    };
  }
}

// --- オプショナル: テスト用関数 ---
function test_FolderStructureAndSheetReadAndHistory() {
  try {
    createHistorySheetIfNotExists();
    Logger.log("履歴シートの存在確認/作成テスト完了。");

    const parentFolderId = getOrCreateFolderIdByName(PARENT_FOLDER_NAME);
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    Logger.log(`親フォルダ: "${parentFolder.getName()}" (ID: ${parentFolderId}) URL: ${parentFolder.getUrl()}`);

    const subFolderName = getSubFolderNameFromSheet();
    let folderPathText = parentFolder.getName();
    let folderUrl = parentFolder.getUrl();

    if (subFolderName) {
      const subFolderId = getOrCreateFolderIdByName(subFolderName, parentFolder);
      const subFolder = DriveApp.getFolderById(subFolderId);
      folderPathText += ` > ${subFolder.getName()}`;
      folderUrl = subFolder.getUrl();
      Logger.log(`サブフォルダ: "${subFolder.getName()}" (ID: ${subFolderId}) URL: ${subFolder.getUrl()}`);
    } else {
      Logger.log(`サブフォルダ名がスプレッドシートから取得できませんでした。親フォルダに記録されます。`);
    }
    Logger.log(`想定保存パス: ${folderPathText}`);

    const testFileName = "test_recording_with_links.mp3";
    const testFileUrl = "https://drive.google.com/file/d/dummy_file_id/view?usp=sharing";
    addRecordToHistorySheet(testFileName, folderPathText, folderUrl, testFileUrl);
    Logger.log("ダミー履歴の記録テスト完了。履歴シートを確認してください (フォルダパスとファイルリンクがハイパーリンクになっているはずです)。");

  } catch (e) {
    Logger.log(`テスト失敗: ${e.toString()}`);
  }
}
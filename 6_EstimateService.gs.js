/**
 * 6_EstimateService.gs v1.2
 * Gemini API版 AI自動読み取り機能
 */

/**
 * 見積IDを生成
 */
function generateEstimateId() {
  const now = new Date();
  const year = now.getFullYear();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMddHHmmss');
  return 'EST-' + year + '-' + timestamp;
}

/**
 * 見積ヘッダーを登録
 */
function saveEstimateHeader(estimateData) {
  const config = getConfig();
  const headerSheet = getSheet(config.SHEET_NAMES.ESTIMATE_HEADER);
  
  const estimateId = generateEstimateId();
  
  const newRow = [
    estimateId,
    estimateData.projectId,
    estimateData.locCode,
    estimateData.locName,
    estimateData.eqId,
    estimateData.eqName,
    estimateData.vendor,
    estimateData.estimateDate,
    estimateData.amountExcludingTax,
    estimateData.consumptionTax,
    estimateData.totalAmount,
    estimateData.expenses,
    estimateData.pdfFileName,
    estimateData.pdfLink,
    new Date()
  ];
  
  headerSheet.appendRow(newRow);
  
  return estimateId;
}

/**
 * 見積明細を登録
 */
function saveEstimateDetails(estimateId, details) {
  const config = getConfig();
  const detailSheet = getSheet(config.SHEET_NAMES.ESTIMATE_DETAIL);
  
  details.forEach(function(item, index) {
    const newRow = [
      estimateId,
      index + 1,
      item.itemName,
      item.unitPrice || 0,
      item.quantity || 0,
      item.unit || '',
      item.subtotal || 0,
      item.note || ''
    ];
    detailSheet.appendRow(newRow);
  });
}

/**
 * 案件の見積一覧を取得
 */
function getEstimatesByProject(projectId) {
  const config = getConfig();
  const headerSheet = getSheet(config.SHEET_NAMES.ESTIMATE_HEADER);
  const detailSheet = getSheet(config.SHEET_NAMES.ESTIMATE_DETAIL);
  
  const headerData = headerSheet.getDataRange().getValues();
  const detailData = detailSheet.getDataRange().getValues();
  
  if (headerData.length <= 1) return [];
  
  const col = {};
  headerData[0].forEach(function(h, i) { col[h] = i; });
  
  const estimates = [];
  
  for (var i = 1; i < headerData.length; i++) {
    const row = headerData[i];
    if (row[col['案件ID']] === projectId) {
      const estimateId = row[col['見積ID']];
      
      const details = [];
      for (var j = 1; j < detailData.length; j++) {
        if (detailData[j][0] === estimateId) {
          details.push({
            rowNumber: detailData[j][1],
            itemName: detailData[j][2],
            unitPrice: detailData[j][3],
            quantity: detailData[j][4],
            unit: detailData[j][5],
            subtotal: detailData[j][6],
            note: detailData[j][7]
          });
        }
      }
      
      estimates.push({
        estimateId: estimateId,
        vendor: row[col['業者名']],
        estimateDate: row[col['見積日']],
        amountExcludingTax: row[col['総額(税抜)']],
        consumptionTax: row[col['消費税']],
        totalAmount: row[col['総額(税込)']],
        expenses: row[col['諸経費']],
        pdfFileName: row[col['PDFファイル名']],
        pdfLink: row[col['PDFリンク']],
        registeredDate: row[col['登録日']],
        details: details
      });
    }
  }
  
  return estimates;
}

/**
 * PDFから見積情報を自動抽出（Gemini API使用）
 */
function extractEstimateFromPDF(pdfFileId) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!apiKey) {
      return {
        success: false,
        message: 'GEMINI_API_KEYが設定されていません'
      };
    }
    
    const file = DriveApp.getFileById(pdfFileId);
    const blob = file.getBlob();
    const base64Data = Utilities.base64Encode(blob.getBytes());
    
    const prompt = `以下の見積書PDFから、情報を抽出してJSON形式で返してください。

必須項目:
- 業者名（会社名）
- 見積日（YYYY-MM-DD形式）
- 総額(税抜)（数値のみ）
- 消費税（数値のみ）
- 総額(税込)（数値のみ）
- 諸経費（数値のみ、なければ0）
- 明細（配列形式で以下を含む）
  - 項目名
  - 単価（数値のみ）
  - 数量（数値のみ）
  - 単位（例: 台、式、日、m2など）
  - 小計（数値のみ）
  - 備考（あれば）

JSONフォーマット:
{
  "vendor": "株式会社○○",
  "estimateDate": "2026-01-03",
  "amountExcludingTax": 1200000,
  "consumptionTax": 120000,
  "totalAmount": 1320000,
  "expenses": 100000,
  "details": [
    {
      "itemName": "計量機本体",
      "unitPrice": 500000,
      "quantity": 4,
      "unit": "台",
      "subtotal": 2000000,
      "note": ""
    }
  ]
}

注意事項:
- 金額はカンマを除いた数値のみ
- 明細は主要な項目のみ（細かい項目は統合可）
- 単価が不明な場合は小計を数量で割る
- 諸経費は交通費、運搬費、管理費などの合計
- JSONのみを返し、説明文は不要
- マークダウンのコードブロック記号は付けない`;

    const payload = {
      contents: [{
        parts: [
          {
            text: prompt
          },
          {
            inline_data: {
              mime_type: 'application/pdf',
              data: base64Data
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 8192
      }
    };
    
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=' + apiKey;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const result = JSON.parse(response.getContentText());
    
    if (result.candidates && result.candidates[0] && result.candidates[0].content) {
      const text = result.candidates[0].content.parts[0].text;
      const jsonText = text.replace(/```json\n?/g, '').replace(/```\n?/g, '').trim();
      const extracted = JSON.parse(jsonText);
      
      return {
        success: true,
        data: extracted
      };
    } else {
      Logger.log('Gemini APIエラー: ' + JSON.stringify(result));
      return {
        success: false,
        message: 'PDFの解析に失敗しました'
      };
    }
  } catch (e) {
    Logger.log('PDF抽出エラー: ' + e.message);
    return {
      success: false,
      message: 'エラー: ' + e.message
    };
  }
}

/**
 * 見積データを保存（ヘッダー + 明細）
 */
function saveEstimate(estimateData, details) {
  try {
    const estimateId = saveEstimateHeader(estimateData);
    saveEstimateDetails(estimateId, details);
    
    return {
      success: true,
      estimateId: estimateId,
      message: '見積を登録しました'
    };
  } catch (e) {
    return {
      success: false,
      message: 'エラー: ' + e.message
    };
  }
}

/**
 * テスト用: 手動でPDF解析を実行
 */
function testExtractPDF() {
  // テスト用のPDFファイルIDを指定
  const testFileId = '1ZkGLJHHe14YKFAGTOHC1JrSNSI3q0xMq';
  
  const result = extractEstimateFromPDF(testFileId);
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * 利用可能なGeminiモデルのリストを取得
 */
function listAvailableModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  const url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + apiKey;
  
  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true
  });
  
  const result = JSON.parse(response.getContentText());
  
  Logger.log('=== 利用可能なモデル一覧 ===');
  
  if (result.models) {
    result.models.forEach(function(model) {
      Logger.log('モデル名: ' + model.name);
      Logger.log('表示名: ' + model.displayName);
      Logger.log('サポート: ' + JSON.stringify(model.supportedGenerationMethods));
      Logger.log('---');
    });
  } else {
    Logger.log('エラー: ' + JSON.stringify(result));
  }
}

/**
 * Gemini APIキーを設定する（初回のみ実行）
 */
function setGeminiApiKey() {
  const apiKey = 'AIzaSyB7vHNMb7N1oK91oHafav-Cbxbfe8mXITw'; // ここに取得したAPIキーを貼り付け
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', apiKey);
  Logger.log('✅ APIキーを設定しました');
}

/**
 * APIキーが正しく設定されているか確認
 */
function checkGeminiApiKey() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (apiKey) {
    Logger.log('✅ APIキーは設定されています');
    Logger.log('キーの先頭: ' + apiKey.substring(0, 20) + '...');
  } else {
    Logger.log('❌ APIキーが設定されていません');
  }
}
/**
 * 見積PDF抽出のテスト
 */
function testEstimateSystem() {
  Logger.log('=== 見積システムのテスト開始 ===');
  
  // 1. 受信BOXのスキャン
  Logger.log('\n【ステップ1】受信BOXをスキャン中...');
  const inboxResult = scanInboxFiles();
  Logger.log('検出ファイル数: ' + inboxResult.files.length);
  
  if (inboxResult.files.length === 0) {
    Logger.log('❌ PDFファイルが見つかりません');
    return;
  }
  
  // 2. 最初のファイルで抽出テスト
  const testFile = inboxResult.files[0];
  Logger.log('\n【ステップ2】PDF抽出テスト');
  Logger.log('テストファイル: ' + testFile.name);
  Logger.log('ファイルID: ' + testFile.id);
  
  const extracted = extractEstimateFromPDF(testFile.id);
  
  if (extracted.success) {
    Logger.log('\n✅ 抽出成功！');
    Logger.log('業者名: ' + extracted.data.vendor);
    Logger.log('見積日: ' + extracted.data.estimateDate);
    Logger.log('総額(税込): ' + extracted.data.totalAmount + '円');
    Logger.log('明細件数: ' + extracted.data.details.length + '件');
    
    // 明細の最初の3件を表示
    Logger.log('\n【明細サンプル】');
    extracted.data.details.slice(0, 3).forEach((item, idx) => {
      Logger.log(`${idx + 1}. ${item.itemName} - ${item.subtotal}円`);
    });
  } else {
    Logger.log('\n❌ 抽出失敗');
    Logger.log('エラー: ' + extracted.message);
  }
  
  Logger.log('\n=== テスト完了 ===');
}

/**
 * スプレッドシート保存テスト（修正版）
 */
function testSaveEstimate() {
  Logger.log('=== スプレッドシート保存テスト ===');
  
  // 1. PDFファイル取得
  const folderInfo = ensureInboxFolder();
  const folder = DriveApp.getFolderById(folderInfo.id);
  const pdfFiles = folder.getFilesByType(MimeType.PDF);
  
  if (!pdfFiles.hasNext()) {
    Logger.log('❌ PDFファイルがありません');
    return;
  }
  
  const testFile = pdfFiles.next();
  Logger.log('テストファイル: ' + testFile.getName());
  
  // 2. PDF抽出
  const result = extractEstimateFromPDF(testFile.getId(), testFile.getName());
  
  if (!result || !result.success) {
    Logger.log('❌ 抽出失敗');
    return;
  }
  
  const extractedData = result.data; // ← ここで data を取り出す
  
  Logger.log('✅ 抽出成功');
  Logger.log('業者名: ' + extractedData.vendor);
  Logger.log('総額: ' + extractedData.totalAmount + '円');
  Logger.log('明細件数: ' + extractedData.details.length + '件');
  
  // 3. スプレッドシート保存
  try {
    const fileInfo = {
      id: testFile.getId(),
      name: testFile.getName(),
      url: testFile.getUrl()
    };
    
    const estimateId = saveEstimateToSheet(result, fileInfo); // result全体を渡す
    Logger.log('\n✅ 保存成功！');
    Logger.log('見積ID: ' + estimateId);
    
    // 4. 保存内容確認
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const compareSheet = ss.getSheetByName('見積比較');
    const lastRow = compareSheet.getLastRow();
    const savedData = compareSheet.getRange(lastRow, 1, 1, 14).getValues()[0];
    
    Logger.log('\n【保存された内容】');
    Logger.log('見積ID: ' + savedData[0]);
    Logger.log('拠点名: ' + savedData[3]);
    Logger.log('設備名: ' + savedData[5]);
    Logger.log('業者名: ' + savedData[6]);
    Logger.log('総額(税込): ' + savedData[10]);
    
  } catch (error) {
    Logger.log('❌ 保存エラー: ' + error.message);
    Logger.log(error.stack);
  }
}

/**
 * 抽出した見積データをスプレッドシートに保存
 * @param {Object} result - extractEstimateFromPDF()の戻り値 {success, data}
 * @param {Object} fileInfo - {id, name, url}
 * @return {string} 見積ID
 */
function saveEstimateToSheet(result, fileInfo) {
  if (!result || !result.success) {
    throw new Error('抽出データが不正です');
  }
  
  const data = result.data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 見積IDを生成（E + タイムスタンプ）
  const estimateId = 'E' + new Date().getTime();
  
  // ファイル名から案件情報を推測
  const activeProjects = getActiveProjects();
  const allEquipments = getEquipmentListCached();
  const suggestion = suggestProjectFromFileName(fileInfo.name, activeProjects, allEquipments);
  
  // suggestionの値を取得（nullの場合は空の値を使用）
  const projectId = suggestion ? suggestion.id || '' : '';
  const locationCode = suggestion ? suggestion.locCode || '' : '';
  const locationName = suggestion ? suggestion.locName || '' : '';
  const equipmentId = suggestion ? suggestion.eqId || '' : '';
  const equipmentName = suggestion ? suggestion.eqName || '' : '';
  
  // 見積比較シートに保存
  const compareSheet = ss.getSheetByName('見積比較');
  const compareRow = [
    estimateId,                           // 見積ID
    projectId,                            // 案件ID
    locationCode,                         // 拠点コード
    locationName,                         // 拠点名
    equipmentId,                          // 設備ID
    equipmentName,                        // 設備名
    data.vendor || '',                    // 業者名
    data.estimateDate || '',              // 見積日
    data.amountExcludingTax || 0,         // 総額(税抜)
    data.consumptionTax || 0,             // 消費税
    data.totalAmount || 0,                // 総額(税込)
    data.expenses || 0,                   // 諸経費
    fileInfo.name,                        // PDFファイル名
    fileInfo.url,                         // PDFリンク
    new Date()                            // 登録日
  ];
  compareSheet.appendRow(compareRow);
  
  // 見積明細シートに保存
  const detailSheet = ss.getSheetByName('見積明細');
  if (data.details && data.details.length > 0) {
    const detailRows = data.details.map((item, index) => [
      estimateId,                         // 見積ID
      index + 1,                          // 行番号
      item.itemName || '',                // 項目名
      item.unitPrice || 0,                // 単価
      item.quantity || 0,                 // 数量
      item.unit || '',                    // 単位
      item.subtotal || 0,                 // 小計
      item.note || ''                     // 備考
    ]);
    
    detailRows.forEach(row => detailSheet.appendRow(row));
  }
  
  Logger.log('✅ スプレッドシートに保存完了: ' + estimateId);
  return estimateId;
}

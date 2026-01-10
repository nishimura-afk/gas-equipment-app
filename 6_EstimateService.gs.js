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
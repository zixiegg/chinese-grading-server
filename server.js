const express = require('express');
const cors = require('cors');
const multer = require('multer');
require('dotenv').config();

const app = express();
const upload = multer({ 
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 }
});

app.use(cors({
  origin: '*',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));

app.use(express.json({ limit: '50mb' }));

// 健康檢查
app.get('/', (req, res) => {
  res.json({ 
    status: 'OK', 
    message: '中文科批改系統 API 服務器運行中',
    timestamp: new Date().toISOString()
  });
});

// 測試 API 連接
app.post('/api/test', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    if (apiType === 'gemini') {
      const result = await testGeminiConnection(apiKey, model);
      return res.json(result);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ success: false, message: '自定義 API 需要提供 API 基礎 URL' });
      }
      const result = await testCustomConnection(apiKey, baseURL, model);
      return res.json(result);
    } else if (apiType === 'openai') {
      const result = await testOpenAIConnection(apiKey, model);
      return res.json(result);
    } else {
      return res.status(400).json({ success: false, message: '未知的 API 類型: ' + apiType });
    }
  } catch (error) {
    console.error('Test API error:', error);
    res.status(500).json({ success: false, message: error.message || '測試失敗' });
  }
});

// OCR 提取文字 - 支持多篇文章分篇
app.post('/api/extract', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, fileType, fileData, text } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    if (!fileData && !text) {
      return res.status(400).json({ success: false, message: '沒有提供文件或文字' });
    }

    console.log('Extract request:', { apiType, fileType, hasFileData: !!fileData, hasText: !!text });

    let result;
    if (apiType === 'gemini') {
      result = await extractWithGemini(apiKey, model, fileData, text, fileType);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ success: false, message: '自定義 API 需要提供 API 基礎 URL' });
      }
      result = await extractWithCustom(apiKey, baseURL, model, fileData, text, fileType);
    } else if (apiType === 'openai') {
      result = await extractWithOpenAI(apiKey, model, fileData, text, fileType);
    } else {
      return res.status(400).json({ success: false, message: '未知的 API 類型: ' + apiType });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Extract error:', error);
    res.status(500).json({ success: false, message: error.message || '提取失敗' });
  }
});

// 提取題目與評分準則
app.post('/api/extract-question-criteria', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, fileType, fileData, text } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    if (!fileData && !text) {
      return res.status(400).json({ success: false, message: '沒有提供文件或文字' });
    }

    console.log('Extract question/criteria request:', { apiType, fileType, hasFileData: !!fileData, hasText: !!text });

    let result;
    if (apiType === 'gemini') {
      result = await extractQuestionCriteriaWithGemini(apiKey, model, fileData, text, fileType);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ success: false, message: '自定義 API 需要提供 API 基礎 URL' });
      }
      result = await extractQuestionCriteriaWithCustom(apiKey, baseURL, model, fileData, text, fileType);
    } else if (apiType === 'openai') {
      result = await extractQuestionCriteriaWithOpenAI(apiKey, model, fileData, text, fileType);
    } else {
      return res.status(400).json({ success: false, message: '未知的 API 類型: ' + apiType });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Extract question/criteria error:', error);
    res.status(500).json({ success: false, message: error.message || '提取失敗' });
  }
});

// 批改作文
app.post('/api/grade', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, essayText, question, customCriteria, gradingMode, contentPriority, enhancementDirection } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    if (!essayText) {
      return res.status(400).json({ success: false, message: '作文內容不能為空' });
    }

    const maxLength = 15000;
    const truncatedText = essayText.length > maxLength 
      ? essayText.substring(0, maxLength) + '\n\n[文章過長，已截斷]' 
      : essayText;

    let result;
    if (apiType === 'gemini') {
      result = await gradeWithGemini(apiKey, model, truncatedText, question, customCriteria, gradingMode, contentPriority, enhancementDirection);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ success: false, message: '自定義 API 需要提供 API 基礎 URL' });
      }
      result = await gradeWithCustom(apiKey, baseURL, model, truncatedText, question, customCriteria, gradingMode, contentPriority, enhancementDirection);
    } else if (apiType === 'openai') {
      result = await gradeWithOpenAI(apiKey, model, truncatedText, question, customCriteria, gradingMode, contentPriority, enhancementDirection);
    } else {
      return res.status(400).json({ success: false, message: '未知的 API 類型: ' + apiType });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Grade error:', error);
    res.status(500).json({ success: false, message: error.message || '批改失敗' });
  }
});

// 生成實用寫作模擬卷（新邏輯：上傳模擬卷→生成新模擬卷）
app.post('/api/generate-practical-exam', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, fileData, fileType, text, genre } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    if (!fileData && !text) {
      return res.status(400).json({ success: false, message: '請上傳模擬卷文件' });
    }

    if (!genre) {
      return res.status(400).json({ success: false, message: '請選擇文體' });
    }

    let result;
    if (apiType === 'gemini') {
      result = await generatePracticalExamWithGemini(apiKey, model, fileData, text, fileType, genre);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ success: false, message: '自定義 API 需要提供 API 基礎 URL' });
      }
      result = await generatePracticalExamWithCustom(apiKey, baseURL, model, fileData, text, fileType, genre);
    } else if (apiType === 'openai') {
      result = await generatePracticalExamWithOpenAI(apiKey, model, fileData, text, fileType, genre);
    } else {
      return res.status(400).json({ success: false, message: '未知的 API 類型: ' + apiType });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Generate exam error:', error);
    res.status(500).json({ success: false, message: error.message || '生成失敗' });
  }
});

// 全班分析
app.post('/api/analyze-class', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, reports, question, gradingMode } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    const maxReports = 30;
    const limitedReports = reports.slice(0, maxReports);

    let result;
    if (apiType === 'gemini') {
      result = await analyzeClassWithGemini(apiKey, model, limitedReports, question, gradingMode);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ success: false, message: '自定義 API 需要提供 API 基礎 URL' });
      }
      result = await analyzeClassWithCustom(apiKey, baseURL, model, limitedReports, question, gradingMode);
    } else if (apiType === 'openai') {
      result = await analyzeClassWithOpenAI(apiKey, model, limitedReports, question, gradingMode);
    } else {
      return res.status(400).json({ success: false, message: '未知的 API 類型: ' + apiType });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Analyze class error:', error);
    res.status(500).json({ success: false, message: error.message || '分析失敗' });
  }
});

// ============ 輔助函數 ============

function normalizeGeminiModelName(modelName) {
  if (!modelName) return 'gemini-2.0-flash';
  let name = modelName.startsWith('models/') ? modelName.substring(7) : modelName;
  const modelMap = {
    'gemini-1.5-flash': 'gemini-1.5-flash-latest',
    'gemini-1.5-pro': 'gemini-1.5-pro-latest',
    'gemini-1.0-pro': 'gemini-1.0-pro-latest',
    'gemini-pro': 'gemini-1.0-pro-latest',
  };
  if (modelMap[name]) return modelMap[name];
  if (name.includes('-latest') || /-\d{3}$/.test(name)) return name;
  return name;
}

// Gemini 安全設置 - 解除過濾
const GEMINI_SAFETY_SETTINGS = [
  { category: 'HARM_CATEGORY_HARASSMENT', threshold: 'BLOCK_NONE' },
  { category: 'HARM_CATEGORY_HATE_SPEECH', threshold: 'BLOCK_NONE' },
  { category: 'HARM_CATEGORY_SEXUALLY_EXPLICIT', threshold: 'BLOCK_NONE' },
  { category: 'HARM_CATEGORY_DANGEROUS_CONTENT', threshold: 'BLOCK_NONE' }
];

// 安全解析 JSON - 增強版
function safeJSONParse(text, context = '') {
  if (!text || typeof text !== 'string') {
    throw new Error(`AI 返回內容為空或格式錯誤 (${context})`);
  }
  
  console.log(`Parsing JSON (${context}):`, text.substring(0, 200) + '...');
  
  try {
    // 1. 去除 markdown 代碼塊標記
    let cleaned = text
      .replace(/```json\s*/gi, '')
      .replace(/```\s*$/gi, '')
      .trim();
    
    // 2. 嘗試直接解析
    try {
      return JSON.parse(cleaned);
    } catch (e) {
      // 3. 嘗試提取 JSON 對象（找第一個 { 和最後一個 }）
      const firstBrace = cleaned.indexOf('{');
      const lastBrace = cleaned.lastIndexOf('}');
      
      if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
        const jsonStr = cleaned.substring(firstBrace, lastBrace + 1);
        try {
          return JSON.parse(jsonStr);
        } catch (e2) {
          console.log('Extracted JSON parse failed, trying repair...');
        }
      }
      
      // 4. 嘗試修復常見問題：去除尾部逗號
      cleaned = cleaned.replace(/,\s*([}\]])/g, '$1');
      
      // 5. 嘗試修復未閉合的字符串
      const openQuotes = (cleaned.match(/"/g) || []).length;
      if (openQuotes % 2 !== 0) {
        cleaned += '"';
      }
      
      try {
        return JSON.parse(cleaned);
      } catch (e3) {
        throw new Error(`無法解析 JSON: ${text.substring(0, 100)}...`);
      }
    }
  } catch (error) {
    console.error('JSON parse error:', error, 'Original text:', text.substring(0, 500));
    throw new Error(`無法解析 AI 返回的 JSON 格式: ${error.message}`);
  }
}

// ============ Gemini API 函數 ============

async function testGeminiConnection(apiKey, modelName = 'gemini-2.0-flash') {
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  try {
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}&pageSize=5`,
      { method: 'GET', headers: { 'Content-Type': 'application/json' } }
    );

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || `HTTP ${response.status}`);
    }

    const data = await response.json();
    const availableModels = data.models?.map(m => m.name) || [];
    const isModelAvailable = availableModels.some(m => m === model || m.endsWith(normalizedModelName));

    return {
      success: true,
      message: isModelAvailable 
        ? `Gemini API 連接成功！模型 "${normalizedModelName}" 可用`
        : `API 連接成功！但模型 "${normalizedModelName}" 可能不可用`,
      model: normalizedModelName
    };
  } catch (error) {
    return handleGeminiError(error, normalizedModelName);
  }
}

async function extractWithGemini(apiKey, modelName, fileData, text, fileType) {
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  console.log('extractWithGemini:', { model, fileType, hasFileData: !!fileData, hasText: !!text });
  
  const prompt = `你是一個專業的OCR文字識別助手。請從圖片或文檔中提取學生的作文文字。

【重要】分辨文章數量的規則：
- 如果文件中只有一篇學生作文，請只返回一篇文章
- 如果文件中有多篇學生作文（例如多個學生的作品、多頁不同學生的作文），請識別並分開每一篇
- 判斷標準：不同學生姓名、明顯的分隔線、不同的標題通常表示不同文章
- 如果只有一個學生姓名且內容連貫，請視為一篇文章

對於每一篇，請提取：
1. 學生姓名（通常在文章開頭或標題處）
2. 學生學號/班級（如果有）
3. 作文正文內容

提取要求：
- 保持原文的段落格式
- 只提取作文正文，不要包含題目（除非題目是文章的一部分）
- 保持所有標點符號
- 不要修改任何文字，包括錯別字

請以JSON格式返回，如果有多篇文章，請返回數組：
{
  "articles": [
    {
      "text": "第一篇作文全文",
      "name": "學生姓名1",
      "studentId": "學號1"
    }
  ]
}

如果只有一篇文章，也使用相同的格式返回。`;

  let requestBody;
  const isImage = fileType && fileType.startsWith('image/');
  const isPDF = fileType && (fileType === 'application/pdf' || fileType.includes('pdf'));
  
  console.log('Type check:', { isImage, isPDF, fileType });

  if (fileData && (isImage || isPDF)) {
    const mimeType = isPDF ? 'application/pdf' : fileType;
    console.log('Using inline_data mode with mimeType:', mimeType);
    
    requestBody = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: fileData
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 65536,
        responseMimeType: 'application/json'
      },
      safetySettings: GEMINI_SAFETY_SETTINGS
    };
  } else {
    console.log('Using text mode');
    const content = text || '';
    requestBody = {
      contents: [{
        parts: [{
          text: prompt + '\n\n請提取以下文本中的學生作文內容：\n\n' + content
        }]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 65536,
        responseMimeType: 'application/json'
      },
      safetySettings: GEMINI_SAFETY_SETTINGS
    };
  }

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody)
    }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'Gemini API 請求失敗');
  }

  const data = await response.json();
  
  // 檢查是否有內容被阻擋
  if (data.promptFeedback?.blockReason) {
    throw new Error(`內容被阻擋: ${data.promptFeedback.blockReason}`);
  }
  
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    // 檢查 finishReason
    const finishReason = data.candidates?.[0]?.finishReason;
    if (finishReason && finishReason !== 'STOP') {
      throw new Error(`Gemini API 返回異常: ${finishReason}`);
    }
    throw new Error('Gemini API 返回內容為空');
  }

  const result = safeJSONParse(content, 'extract');
  
  // 支持多篇文章返回
  if (result.articles && Array.isArray(result.articles) && result.articles.length > 0) {
    const firstArticle = result.articles[0];
    return {
      text: firstArticle.text || '',
      name: firstArticle.name || '',
      studentId: firstArticle.studentId || '',
      articles: result.articles
    };
  }
  
  return {
    text: result.text || '',
    name: result.name || '',
    studentId: result.studentId || '',
    articles: result.articles
  };
}

// 提取題目與評分準則
async function extractQuestionCriteriaWithGemini(apiKey, modelName, fileData, text, fileType) {
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  const prompt = `你是一個專業的教育文件分析助手。請從文件中提取作文題目和評分準則。

請分析文件內容，識別並提取：
1. 題目：作文的題目要求
2. 評分準則：如果有提供評分準則，請完整提取；如果沒有，請返回空字符串

請以JSON格式返回：
{
  "question": "提取的題目內容",
  "criteria": "提取的評分準則（如無則為空字符串）"
}`;

  let requestBody;
  const isImage = fileType && fileType.startsWith('image/');
  const isPDF = fileType && (fileType === 'application/pdf' || fileType.includes('pdf'));

  if (fileData && (isImage || isPDF)) {
    const mimeType = isPDF ? 'application/pdf' : fileType;
    requestBody = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: fileData
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 65536,
        responseMimeType: 'application/json'
      },
      safetySettings: GEMINI_SAFETY_SETTINGS
    };
  } else {
    requestBody = {
      contents: [{
        parts: [{
          text: prompt + '\n\n請分析以下內容：\n\n' + (text || '')
        }]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 65536,
        responseMimeType: 'application/json'
      },
      safetySettings: GEMINI_SAFETY_SETTINGS
    };
  }

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody)
    }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'Gemini API 請求失敗');
  }

  const data = await response.json();
  
  if (data.promptFeedback?.blockReason) {
    throw new Error(`內容被阻擋: ${data.promptFeedback.blockReason}`);
  }
  
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    throw new Error('Gemini API 返回內容為空');
  }

  return safeJSONParse(content, 'extract-question-criteria');
}

async function gradeWithGemini(apiKey, modelName, essayText, question, customCriteria, gradingMode = 'secondary', contentPriority = false, enhancementDirection = 'auto') {
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  const systemPrompt = buildGradingPrompt(gradingMode, contentPriority, enhancementDirection);
  const userPrompt = buildUserPrompt(essayText, question, customCriteria, gradingMode);

  const requestBody = {
    contents: [{
      parts: [{ text: systemPrompt + '\n\n' + userPrompt }]
    }],
    generationConfig: {
      temperature: 0.5,
      maxOutputTokens: 65536,
      responseMimeType: 'application/json'
    },
    safetySettings: GEMINI_SAFETY_SETTINGS
  };

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody)
    }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'Gemini API 請求失敗');
  }

  const data = await response.json();
  
  // 檢查是否有內容被阻擋
  if (data.promptFeedback?.blockReason) {
    throw new Error(`內容被阻擋: ${data.promptFeedback.blockReason}`);
  }
  
  let content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    // 嘗試從其他候選人獲取內容
    if (data.candidates && data.candidates.length > 1) {
      for (let i = 1; i < data.candidates.length; i++) {
        const altContent = data.candidates[i]?.content?.parts?.[0]?.text;
        if (altContent) {
          console.log('Using alternative candidate:', i);
          content = altContent;
          break;
        }
      }
    }
    
    if (!content) {
      const finishReason = data.candidates?.[0]?.finishReason;
      throw new Error(`Gemini API 返回內容為空 (finishReason: ${finishReason || 'unknown'})`);
    }
  }

  return parseGradingResult(content, essayText, gradingMode);
}

// 生成實用寫作模擬卷（新邏輯）
async function generatePracticalExamWithGemini(apiKey, modelName, fileData, text, fileType, genre) {
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  const genreNames = {
    speech: '演講辭',
    letter: '書信/公開信',
    proposal: '建議書',
    report: '報告',
    commentary: '評論文章',
    article: '專題文章'
  };

  const prompt = `你是一位專業的DSE中文科試題設計專家。請分析用戶上傳的模擬卷，理解其主題和結構，然後生成一份全新的模擬卷。

【重要要求】
1. 新模擬卷必須保持與原卷相同的主題/主題方向
2. 但內容必須完全不同（不同的情境、不同的資料、不同的具體要求）
3. 用戶選擇的文體是：${genreNames[genre] || genre}
4. 新模擬卷必須符合該文體的格式要求

請以JSON格式返回：
{
  "examPaper": {
    "title": "試卷標題",
    "time": "考試時間（如：45分鐘）",
    "marks": "佔分（如：50分）",
    "instructions": ["考生須知1", "考生須知2"],
    "question": "題目描述",
    "material1": {
      "title": "資料一標題",
      "content": "資料一內容"
    },
    "material2": {
      "title": "資料二標題",
      "content": "資料二內容"
    }
  },
  "markingScheme": {
    "content": {
      "infoPoints": ["內容要點1", "內容要點2"],
      "developmentPoints": ["發展要求1", "發展要求2"]
    },
    "organization": {
      "formatRequirements": ["格式要求1", "格式要求2"],
      "toneRequirements": ["語氣要求1", "語氣要求2"]
    }
  }
}`;

  let requestBody;
  const isImage = fileType && fileType.startsWith('image/');
  const isPDF = fileType && (fileType === 'application/pdf' || fileType.includes('pdf'));

  if (fileData && (isImage || isPDF)) {
    const mimeType = isPDF ? 'application/pdf' : fileType;
    requestBody = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: mimeType,
              data: fileData
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 65536,
        responseMimeType: 'application/json'
      },
      safetySettings: GEMINI_SAFETY_SETTINGS
    };
  } else {
    requestBody = {
      contents: [{
        parts: [{
          text: prompt + '\n\n請根據以下模擬卷內容生成新模擬卷：\n\n' + (text || '')
        }]
      }],
      generationConfig: {
        temperature: 0.7,
        maxOutputTokens: 65536,
        responseMimeType: 'application/json'
      },
      safetySettings: GEMINI_SAFETY_SETTINGS
    };
  }

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody)
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    console.error('Gemini API error response:', errorText);
    throw new Error(`Gemini API 請求失敗: ${response.status}`);
  }

  const data = await response.json();
  
  if (data.promptFeedback?.blockReason) {
    throw new Error(`內容被阻擋: ${data.promptFeedback.blockReason}`);
  }
  
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    throw new Error('Gemini API 返回內容為空');
  }

  return safeJSONParse(content, 'generate-exam');
}

async function analyzeClassWithGemini(apiKey, modelName, reports, question, gradingMode = 'secondary') {
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  const stats = {
    totalStudents: reports.length,
    averageScore: reports.reduce((sum, r) => sum + r.totalScore, 0) / reports.length,
  };

  let systemPrompt;
  if (gradingMode === 'primary') {
    systemPrompt = `你是一位專業的香港小學中文科教師，正在分析全班學生的寫作表現。

請根據學生的評分數據，從以下五個方面進行分析：
1. 切題與內容分析
2. 感受與立意分析
3. 組織與結構分析
4. 語言運用分析
5. 文類與格式分析

最後提供針對性的寫作教學建議。

請以JSON格式返回：
{
  "materialAnalysis": "選材分析內容",
  "relevanceAnalysis": "扣題分析內容",
  "themeAnalysis": "立意分析內容",
  "techniqueAnalysis": "寫作手法分析內容",
  "teachingSuggestion": "寫作教學建議內容"
}`;
  } else if (gradingMode === 'practical') {
    systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的實用寫作表現。

請根據學生的評分數據，從以下方面進行分析：
1. 內容資訊分析
2. 拓展分析
3. 語氣分析
4. 組織結構分析

最後提供針對性的寫作教學建議。

請以JSON格式返回：
{
  "materialAnalysis": "內容資訊分析",
  "relevanceAnalysis": "拓展分析",
  "themeAnalysis": "語氣分析",
  "techniqueAnalysis": "組織結構分析",
  "teachingSuggestion": "寫作教學建議內容"
}`;
  } else {
    systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的寫作表現。

請根據學生的評分數據，從以下四個方面進行分析：
1. 選材分析：學生選取素材的類型、優點和問題
2. 扣題分析：學生對題目的理解程度、常見問題和改進建議
3. 立意分析：學生立意的深度分佈、優點和問題
4. 寫作手法分析：學生使用的寫作手法、優點和問題

最後提供針對性的寫作教學建議。

請以JSON格式返回：
{
  "materialAnalysis": "選材分析內容",
  "relevanceAnalysis": "扣題分析內容",
  "themeAnalysis": "立意分析內容",
  "techniqueAnalysis": "寫作手法分析內容",
  "teachingSuggestion": "寫作教學建議內容"
}`;
  }

  const userPrompt = `請分析以下全班寫作數據：

## 題目
${question}

## 統計數據
- 總人數: ${stats.totalStudents}
- 平均分: ${stats.averageScore.toFixed(1)}

## 學生分數詳情
${reports.map(r => `- ${r.studentWork?.name || '未命名'}: 總分${r.totalScore}`).join('\n')}

請以專業教師的角度進行分析，並以JSON格式返回結果。`;

  const requestBody = {
    contents: [{
      parts: [{ text: systemPrompt + '\n\n' + userPrompt }]
    }],
    generationConfig: {
      temperature: 0.5,
      maxOutputTokens: 65536,
      responseMimeType: 'application/json'
    },
    safetySettings: GEMINI_SAFETY_SETTINGS
  };

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/${model}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody)
    }
  );

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'Gemini API 請求失敗');
  }

  const data = await response.json();
  
  if (data.promptFeedback?.blockReason) {
    throw new Error(`內容被阻擋: ${data.promptFeedback.blockReason}`);
  }
  
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    throw new Error('Gemini API 返回內容為空');
  }

  return safeJSONParse(content, 'analyze-class');
}

// ============ OpenAI API 函數 ============

async function testOpenAIConnection(apiKey, modelName = 'gpt-4o') {
  try {
    const response = await fetch('https://api.openai.com/v1/models', {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${apiKey}` }
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || `HTTP ${response.status}`);
    }

    const data = await response.json();
    const availableModels = data.data?.map(m => m.id) || [];
    const isModelAvailable = availableModels.some(m => m === modelName);

    return {
      success: true,
      message: isModelAvailable 
        ? `OpenAI API 連接成功！模型 "${modelName}" 可用`
        : `API 連接成功！但模型 "${modelName}" 可能不可用`,
      model: modelName
    };
  } catch (error) {
    return handleOpenAIError(error, modelName);
  }
}

async function extractWithOpenAI(apiKey, modelName, fileData, text, fileType) {
  const model = modelName || 'gpt-4o';
  
  const systemPrompt = `你是一個專業的OCR文字識別助手。請從圖片或文檔中提取學生的作文文字。

【重要】分辨文章數量的規則：
- 如果文件中只有一篇學生作文，請只返回一篇文章
- 如果文件中有多篇學生作文，請識別並分開每一篇
- 判斷標準：不同學生姓名、明顯的分隔線、不同的標題通常表示不同文章

提取要求：
1. 保持原文的段落格式
2. 識別學生姓名和學號（通常在文章開頭或標題處）
3. 只提取作文正文，不要包含題目（除非題目是文章的一部分）
4. 保持所有標點符號
5. 不要修改任何文字，包括錯別字

請以JSON格式返回：
{
  "articles": [
    {
      "text": "作文全文",
      "name": "學生姓名（如無則留空）",
      "studentId": "學生學號（如無則留空）"
    }
  ]
}`;

  let messages;
  
  if (fileData && fileType && fileType.startsWith('image/')) {
    messages = [
      { role: 'system', content: systemPrompt },
      {
        role: 'user',
        content: [
          { type: 'text', text: '請提取這張圖片中的學生作文文字，並識別姓名和學號。' },
          { type: 'image_url', image_url: { url: `data:${fileType};base64,${fileData}` } }
        ]
      }
    ];
  } else {
    const content = text || '';
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: `請提取以下文本中的學生作文內容：\n\n${content}` }
    ];
  }

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.3,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'OpenAI API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('OpenAI API 返回內容為空');
  }

  const result = JSON.parse(content);
  
  if (result.articles && Array.isArray(result.articles) && result.articles.length > 0) {
    const firstArticle = result.articles[0];
    return {
      text: firstArticle.text || '',
      name: firstArticle.name || '',
      studentId: firstArticle.studentId || '',
      articles: result.articles
    };
  }
  
  return {
    text: result.text || '',
    name: result.name || '',
    studentId: result.studentId || ''
  };
}

async function extractQuestionCriteriaWithOpenAI(apiKey, modelName, fileData, text, fileType) {
  const model = modelName || 'gpt-4o';
  
  const systemPrompt = `你是一個專業的教育文件分析助手。請從文件中提取作文題目和評分準則。

請分析文件內容，識別並提取：
1. 題目：作文的題目要求
2. 評分準則：如果有提供評分準則，請完整提取；如果沒有，請返回空字符串

請以JSON格式返回：
{
  "question": "提取的題目內容",
  "criteria": "提取的評分準則（如無則為空字符串）"
}`;

  let messages;
  
  if (fileData && fileType && fileType.startsWith('image/')) {
    messages = [
      { role: 'system', content: systemPrompt },
      {
        role: 'user',
        content: [
          { type: 'text', text: '請分析這張圖片，提取題目和評分準則。' },
          { type: 'image_url', image_url: { url: `data:${fileType};base64,${fileData}` } }
        ]
      }
    ];
  } else {
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: `請分析以下內容，提取題目和評分準則：\n\n${text || ''}` }
    ];
  }

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.3,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'OpenAI API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('OpenAI API 返回內容為空');
  }

  return JSON.parse(content);
}

async function gradeWithOpenAI(apiKey, modelName, essayText, question, customCriteria, gradingMode = 'secondary', contentPriority = false, enhancementDirection = 'auto') {
  const model = modelName || 'gpt-4o';
  
  const systemPrompt = buildGradingPrompt(gradingMode, contentPriority, enhancementDirection);
  const userPrompt = buildUserPrompt(essayText, question, customCriteria, gradingMode);

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: 0.5,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'OpenAI API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('OpenAI API 返回內容為空');
  }

  return parseGradingResult(content, essayText, gradingMode);
}

async function generatePracticalExamWithOpenAI(apiKey, modelName, fileData, text, fileType, genre) {
  const model = modelName || 'gpt-4o';
  
  const genreNames = {
    speech: '演講辭',
    letter: '書信/公開信',
    proposal: '建議書',
    report: '報告',
    commentary: '評論文章',
    article: '專題文章'
  };

  const systemPrompt = `你是一位專業的DSE中文科試題設計專家。請分析用戶上傳的模擬卷，理解其主題和結構，然後生成一份全新的模擬卷。

【重要要求】
1. 新模擬卷必須保持與原卷相同的主題/主題方向
2. 但內容必須完全不同（不同的情境、不同的資料、不同的具體要求）
3. 用戶選擇的文體是：${genreNames[genre] || genre}
4. 新模擬卷必須符合該文體的格式要求

請以JSON格式返回。`;

  let messages;
  
  if (fileData && fileType && fileType.startsWith('image/')) {
    messages = [
      { role: 'system', content: systemPrompt },
      {
        role: 'user',
        content: [
          { type: 'text', text: '請分析這張模擬卷圖片，理解其主題，然後生成一份全新的模擬卷。' },
          { type: 'image_url', image_url: { url: `data:${fileType};base64,${fileData}` } }
        ]
      }
    ];
  } else {
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: `請根據以下模擬卷內容生成新模擬卷：\n\n${text || ''}` }
    ];
  }

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.7,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'OpenAI API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('OpenAI API 返回內容為空');
  }

  const result = JSON.parse(content);
  return {
    examPaper: result.examPaper || {
      title: 'DSE 中文卷二甲部：實用寫作',
      time: '45分鐘',
      marks: '50分',
      instructions: [],
      question: '',
      material1: { title: '', content: '' },
      material2: { title: '', content: '' },
    },
    markingScheme: result.markingScheme || {
      content: { infoPoints: [], developmentPoints: [] },
      organization: { formatRequirements: [], toneRequirements: [] },
    },
  };
}

async function analyzeClassWithOpenAI(apiKey, modelName, reports, question, gradingMode = 'secondary') {
  const model = modelName || 'gpt-4o';
  
  const stats = {
    totalStudents: reports.length,
    averageScore: reports.reduce((sum, r) => sum + r.totalScore, 0) / reports.length,
  };

  let systemPrompt;
  if (gradingMode === 'primary') {
    systemPrompt = `你是一位專業的香港小學中文科教師，正在分析全班學生的寫作表現。請以JSON格式返回分析結果。`;
  } else if (gradingMode === 'practical') {
    systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的實用寫作表現。請以JSON格式返回分析結果。`;
  } else {
    systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的寫作表現。請以JSON格式返回分析結果。`;
  }

  const userPrompt = `請分析以下全班寫作數據：

## 題目
${question}

## 統計數據
- 總人數: ${stats.totalStudents}
- 平均分: ${stats.averageScore.toFixed(1)}

## 學生分數詳情
${reports.map(r => `- ${r.studentWork?.name || '未命名'}: 總分${r.totalScore}`).join('\n')}

請以專業教師的角度進行分析，並以JSON格式返回結果。`;

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: 0.5,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'OpenAI API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('OpenAI API 返回內容為空');
  }

  return JSON.parse(content);
}

// ============ 自定義 API 函數 ============

async function testCustomConnection(apiKey, baseURL, modelName) {
  try {
    const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
    
    const response = await fetch(`${normalizedURL}/models`, {
      method: 'GET',
      headers: { 'Authorization': `Bearer ${apiKey}` }
    });

    if (!response.ok) {
      const error = await response.json();
      throw new Error(error.error?.message || `HTTP ${response.status}`);
    }

    const data = await response.json();
    const availableModels = data.data?.map(m => m.id) || [];
    const isModelAvailable = availableModels.some(m => m === modelName);

    return {
      success: true,
      message: isModelAvailable 
        ? `API 連接成功！模型 "${modelName}" 可用`
        : `API 連接成功！但模型 "${modelName}" 可能不可用`,
      model: modelName
    };
  } catch (error) {
    return handleOpenAIError(error, modelName);
  }
}

async function extractWithCustom(apiKey, baseURL, modelName, fileData, text, fileType) {
  const model = modelName || 'gpt-4o';
  const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
  
  const systemPrompt = `你是一個專業的OCR文字識別助手。請從圖片或文檔中提取學生的作文文字。

提取要求：
1. 保持原文的段落格式
2. 識別學生姓名和學號（通常在文章開頭或標題處）
3. 只提取作文正文，不要包含題目（除非題目是文章的一部分）
4. 保持所有標點符號
5. 不要修改任何文字，包括錯別字

請以JSON格式返回：
{
  "articles": [
    {
      "text": "作文全文",
      "name": "學生姓名（如無則留空）",
      "studentId": "學生學號（如無則留空）"
    }
  ]
}`;

  let messages;
  
  if (fileData && fileType && fileType.startsWith('image/')) {
    messages = [
      { role: 'system', content: systemPrompt },
      {
        role: 'user',
        content: [
          { type: 'text', text: '請提取這張圖片中的學生作文文字，並識別姓名和學號。' },
          { type: 'image_url', image_url: { url: `data:${fileType};base64,${fileData}` } }
        ]
      }
    ];
  } else {
    const content = text || '';
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: `請提取以下文本中的學生作文內容：\n\n${content}` }
    ];
  }

  const response = await fetch(`${normalizedURL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.3,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('API 返回內容為空');
  }

  const result = JSON.parse(content);
  
  if (result.articles && Array.isArray(result.articles) && result.articles.length > 0) {
    const firstArticle = result.articles[0];
    return {
      text: firstArticle.text || '',
      name: firstArticle.name || '',
      studentId: firstArticle.studentId || '',
      articles: result.articles
    };
  }
  
  return {
    text: result.text || '',
    name: result.name || '',
    studentId: result.studentId || ''
  };
}

async function extractQuestionCriteriaWithCustom(apiKey, baseURL, modelName, fileData, text, fileType) {
  const model = modelName || 'gpt-4o';
  const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
  
  const systemPrompt = `你是一個專業的教育文件分析助手。請從文件中提取作文題目和評分準則。

請以JSON格式返回：
{
  "question": "提取的題目內容",
  "criteria": "提取的評分準則（如無則為空字符串）"
}`;

  let messages;
  
  if (fileData && fileType && fileType.startsWith('image/')) {
    messages = [
      { role: 'system', content: systemPrompt },
      {
        role: 'user',
        content: [
          { type: 'text', text: '請分析這張圖片，提取題目和評分準則。' },
          { type: 'image_url', image_url: { url: `data:${fileType};base64,${fileData}` } }
        ]
      }
    ];
  } else {
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: `請分析以下內容：\n\n${text || ''}` }
    ];
  }

  const response = await fetch(`${normalizedURL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.3,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('API 返回內容為空');
  }

  return JSON.parse(content);
}

async function gradeWithCustom(apiKey, baseURL, modelName, essayText, question, customCriteria, gradingMode = 'secondary', contentPriority = false, enhancementDirection = 'auto') {
  const model = modelName || 'gpt-4o';
  const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
  
  const systemPrompt = buildGradingPrompt(gradingMode, contentPriority, enhancementDirection);
  const userPrompt = buildUserPrompt(essayText, question, customCriteria, gradingMode);

  const response = await fetch(`${normalizedURL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: 0.5,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('API 返回內容為空');
  }

  return parseGradingResult(content, essayText, gradingMode);
}

async function generatePracticalExamWithCustom(apiKey, baseURL, modelName, fileData, text, fileType, genre) {
  const model = modelName || 'gpt-4o';
  const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
  
  const genreNames = {
    speech: '演講辭',
    letter: '書信/公開信',
    proposal: '建議書',
    report: '報告',
    commentary: '評論文章',
    article: '專題文章'
  };

  const systemPrompt = `你是一位專業的DSE中文科試題設計專家。請分析用戶上傳的模擬卷，理解其主題和結構，然後生成一份全新的模擬卷。

用戶選擇的文體是：${genreNames[genre] || genre}

請以JSON格式返回。`;

  let messages;
  
  if (fileData && fileType && fileType.startsWith('image/')) {
    messages = [
      { role: 'system', content: systemPrompt },
      {
        role: 'user',
        content: [
          { type: 'text', text: '請分析這張模擬卷圖片，生成新模擬卷。' },
          { type: 'image_url', image_url: { url: `data:${fileType};base64,${fileData}` } }
        ]
      }
    ];
  } else {
    messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: `請根據以下模擬卷內容生成新模擬卷：\n\n${text || ''}` }
    ];
  }

  const response = await fetch(`${normalizedURL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: messages,
      temperature: 0.7,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('API 返回內容為空');
  }

  const result = JSON.parse(content);
  return {
    examPaper: result.examPaper || {
      title: 'DSE 中文卷二甲部：實用寫作',
      time: '45分鐘',
      marks: '50分',
      instructions: [],
      question: '',
      material1: { title: '', content: '' },
      material2: { title: '', content: '' },
    },
    markingScheme: result.markingScheme || {
      content: { infoPoints: [], developmentPoints: [] },
      organization: { formatRequirements: [], toneRequirements: [] },
    },
  };
}

async function analyzeClassWithCustom(apiKey, baseURL, modelName, reports, question, gradingMode = 'secondary') {
  const model = modelName || 'gpt-4o';
  const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
  
  const stats = {
    totalStudents: reports.length,
    averageScore: reports.reduce((sum, r) => sum + r.totalScore, 0) / reports.length,
  };

  let systemPrompt;
  if (gradingMode === 'primary') {
    systemPrompt = `你是一位專業的香港小學中文科教師，正在分析全班學生的寫作表現。請以JSON格式返回分析結果。`;
  } else if (gradingMode === 'practical') {
    systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的實用寫作表現。請以JSON格式返回分析結果。`;
  } else {
    systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的寫作表現。請以JSON格式返回分析結果。`;
  }

  const userPrompt = `請分析以下全班寫作數據：

## 題目
${question}

## 統計數據
- 總人數: ${stats.totalStudents}
- 平均分: ${stats.averageScore.toFixed(1)}

## 學生分數詳情
${reports.map(r => `- ${r.studentWork?.name || '未命名'}: 總分${r.totalScore}`).join('\n')}

請以專業教師的角度進行分析，並以JSON格式返回結果。`;

  const response = await fetch(`${normalizedURL}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model: model,
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt }
      ],
      temperature: 0.5,
      max_tokens: 4000,
      response_format: { type: 'json_object' }
    })
  });

  if (!response.ok) {
    const error = await response.json();
    throw new Error(error.error?.message || 'API 請求失敗');
  }

  const data = await response.json();
  const content = data.choices[0]?.message?.content;
  
  if (!content) {
    throw new Error('API 返回內容為空');
  }

  return JSON.parse(content);
}

// ============ 錯誤處理函數 ============

function handleGeminiError(error, modelName) {
  const message = error.message || '未知錯誤';
  
  if (message.includes('quota') || message.includes('exceeded') || message.includes('429')) {
    return { success: false, message: 'Gemini API 配額已用完，請檢查您的 Google AI Studio 配額或稍後再試' };
  }
  
  if (message.includes('not found') || message.includes('model')) {
    return { success: false, message: `模型 "${modelName}" 不存在或無權訪問，請確認模型名稱正確` };
  }
  
  if (message.includes('401') || message.includes('Unauthorized') || message.includes('API key not valid')) {
    return { success: false, message: 'API 密鑰無效，請檢查您的 Gemini API 密鑰是否正確' };
  }
  
  if (message.includes('fetch') || message.includes('network') || message.includes('ECONNREFUSED')) {
    return { success: false, message: '無法連接到 Gemini 服務器，請檢查網絡連接' };
  }
  
  if (message.includes('token') || message.includes('Token') || message.includes('1048576')) {
    return { success: false, message: '輸入內容太長，超過了 Token 限制。請嘗試：1) 使用更小的文件 2) 分段處理 3) 使用支持更大上下文的模型' };
  }
  
  return { success: false, message: `連接失敗: ${message}` };
}

function handleOpenAIError(error, modelName) {
  const message = error.message || '未知錯誤';
  
  if (message.includes('quota') || message.includes('exceeded') || message.includes('429')) {
    return { success: false, message: 'API 配額已用完，請檢查您的賬戶配額' };
  }
  
  if (message.includes('not found') || message.includes('model')) {
    return { success: false, message: `模型 "${modelName}" 不存在或無權訪問` };
  }
  
  if (message.includes('401') || message.includes('Unauthorized')) {
    return { success: false, message: 'API 密鑰無效，請檢查您的 API 密鑰' };
  }
  
  if (message.includes('token') || message.includes('Token')) {
    return { success: false, message: '輸入內容太長，超過了 Token 限制。請嘗試：1) 使用更小的文件 2) 分段處理 3) 使用支持更大上下文的模型' };
  }
  
  return { success: false, message: `連接失敗: ${message}` };
}

// ============ Prompt 和解析函數 ============

function buildGradingPrompt(gradingMode = 'secondary', contentPriority = false, enhancementDirection = 'auto') {
  if (gradingMode === 'primary') {
    return buildPrimaryGradingPrompt();
  } else if (gradingMode === 'practical') {
    return buildPracticalGradingPrompt();
  } else {
    return buildSecondaryGradingPrompt(contentPriority, enhancementDirection);
  }
}

function buildSecondaryGradingPrompt(contentPriority = false, enhancementDirection = 'auto') {
  const enhancementInstruction = enhancementDirection === 'narrative' 
    ? '增潤時優先考慮記敘和抒情元素，讓文章更有情感感染力'
    : enhancementDirection === 'argumentative'
    ? '增潤時優先考慮說理和論證元素，讓文章更有說服力'
    : '根據文章類型自動判斷增潤方向';

  const contentPriorityInstruction = contentPriority 
    ? '【以內容為主評分】請優先根據內容質量給分，結構分數可根據內容表現適度調整，但內容與結構分數一般不應相差超過2級' 
    : '【標準評分】請按照HKDSE標準評分準則進行評分';

  return `你是一位專業的香港中學中文科教師，正在批改HKDSE中文卷二乙部命題寫作。

${contentPriorityInstruction}

## 評分準則（必須嚴格遵守）

### 內容（40分）- 品第制 1-10分
評分重點：立意深度、取材恰當度、闡述飽滿度

具體標準：
- 上上(10分): 立意極豐富深刻，取材極恰當，闡述極飽滿。能深入探討題目，觀點獨到，例子精準有力。
- 上中(9分): 立意極豐富尚算深刻，取材極恰當，闡述飽滿。觀點清晰，論證充分。
- 上下(8分): 立意豐富尚算深刻，取材恰當，闡述飽滿。有個人見解，例子適切。
- 中上(7分): 立意平穩具體，取材合理平穩，闡述合理。符合題目要求，內容完整。
- 中中上(6分): 立意一般恰當，取材一般，闡述一般。基本符合題目，但深度不足。
- 中中下(5分): 立意尚具單薄，取材單薄，闡述浮淺。內容簡單，缺乏深度。
- 中下(4分): 立意十分單薄，取材薄弱，闡述極浮淺。內容貧乏。
- 下上(3分): 立意模糊不太相干，取材不扣連，闡述欠奉。偏離題目。
- 下中(2分): 立意十分模糊錯亂，取材混亂，闡述空泛。嚴重離題。
- 下下(1分): 無立意或錯亂極多，取材闡述闕如。完全離題或無法理解。

### 表達（30分）- 品第制 1-10分
評分重點：詞彙運用、文句流暢度、修辭手法

具體標準：
- 上上(10分): 用詞極精確豐富，文句極簡潔流暢，手法純熟靈活。詞藻優美，句式多變，修辭恰當。
- 上中(9分): 用詞極精確豐富，文句極簡潔流暢，手法純熟。
- 上下(8分): 用詞精確豐富，文句簡潔流暢，手法靈活。
- 中上(7分): 用詞準確平穩，文句通順偶有瑕疵，手法尚算靈活。
- 中中上(6分): 用詞大致準確，文句大致通順有沙石，表達一般。
- 中中下(5分): 用詞尚算準確，文句尚算通順有明顯語病，表達浮淺。
- 中下(4分): 用詞粗疏，文句不通順失誤多，表達薄弱。
- 下上(3分): 用詞不準，文句欠通順，表達混亂。
- 下中(2分): 用詞句式嚴重錯誤，表達極混亂。
- 下下(1分): 文句無法達意。

### 結構（20分）- 品第制 1-10分
評分重點：結構完整性、詳略安排、鋪排有序

具體標準：
- 上上(10分): 結構極完整，詳略極得宜，鋪排主次有序。起承轉合自然，段落分明。
- 上中(9分): 結構極完整，詳略得宜，鋪排有序。
- 上下(8分): 結構完整，詳略得宜，鋪排有序。
- 中上(7分): 結構大致完整，詳略大致合宜，過渡尚算自然。
- 中中上(6分): 結構尚具完整，詳略一般有輕微失衡。
- 中中下(5分): 結構尚具，詳略稍失衡，偶有散亂。
- 中下(4分): 尚具組織，詳略明顯失衡，鋪排失當。
- 下上(3分): 組織散亂，詳略嚴重失衡。
- 下中(2分): 毫無組織，鋪排極混亂。
- 下下(1分): 完全無結構。

### 標點（10分）- 5-10分
- 10分: 標點符號運用正確無誤
- 9分: 偶有失誤（1-2處）
- 8分: 有少許失誤（3-4處）
- 7分: 有失誤（5-6處）
- 6分: 失誤較多（7-8處）
- 5分: 有明顯問題（9處以上）

## 重要評分規則（必須遵守）
1. 內容與結構分數一般不應相差超過2級
2. 若內容離題（3分及以下），結構最高只能評至7分
3. 標點分數下限為5分
4. 請根據文章的實際表現給分，不要過度寬鬆或嚴苛

## 增潤文章要求
${enhancementInstruction}
- 增潤後的文章應保持學生的原意和風格
- 修正語病和錯別字
- 提升表達質量，但避免過度堆砌詞藻
- 保持人文情懷，避免AI感

## 輸出格式
請以JSON格式返回，確保所有字段都存在：
{
  "grading": {
    "content": 6,
    "expression": 6,
    "structure": 6,
    "punctuation": 7
  },
  "overallComment": "簡潔易明的總評，2-3段，指出主要優點和可改善之處",
  "contentFeedback": {
    "strengths": ["內容優點1", "內容優點2"],
    "improvements": ["內容改善建議1", "內容改善建議2"]
  },
  "expressionFeedback": {
    "strengths": ["表達優點1", "表達優點2"],
    "improvements": ["表達改善建議1", "表達改善建議2"]
  },
  "structureFeedback": {
    "strengths": ["結構優點1", "結構優點2"],
    "improvements": ["結構改善建議1", "結構改善建議2"]
  },
  "punctuationFeedback": {
    "strengths": ["標點優點"],
    "improvements": ["標點改善建議"]
  },
  "enhancedText": "增潤後的完整文章，避免AI堆砌感，要有人文情懷",
  "enhancementNotes": ["修改說明1：具體說明修改了什麼", "修改說明2"],
  "modelEssay": "一篇符合HKDSE上上品標準的奪星文章示範，展示如何更好地處理這個題目"
}`;
}

function buildPrimaryGradingPrompt() {
  return `你是一位專業的香港小學中文科教師，正在批改小學命題寫作。

## 評分準則（必須嚴格遵守）

### 切題與內容（30分）- 等級制 1-4級
評分重點：切合題旨、內容具體、素材運用

具體標準：
- 4級(24-30分): 內容切題，具體充實，能圍繞主題展開，有適當的例子和細節
- 3級(18-23分): 內容大致切題，尚算具體，能圍繞主題但深度不足
- 2級(9-17分): 內容尚切題但不夠具體，有離題或空洞的情況
- 1級(0-8分): 內容不切題或過於簡略，嚴重離題或內容貧乏

### 感受與立意（20分）- 等級制 1-4級
評分重點：情感真摯、立意清晰、個人體會

具體標準：
- 4級(16-20分): 感受真摯深刻，立意清晰，有個人獨特體會
- 3級(12-15分): 感受真摯，立意尚清晰，有個人體會
- 2級(6-11分): 感受一般，立意不夠清晰，體會較淺
- 1級(0-5分): 感受虛假或缺乏，立意模糊，沒有個人體會

### 組織與結構（20分）- 等級制 1-4級
評分重點：段落分明、详略得當、過渡自然

具體標準：
- 4級(16-20分): 結構完整，段落分明，详略得當，過渡自然
- 3級(12-15分): 結構尚完整，段落尚分明，详略大致得當
- 2級(6-11分): 結構一般，段落不夠分明，详略有失衡
- 1級(0-5分): 結構散亂，段落不清，详略嚴重失衡

### 語言運用（20分）- 等級制 1-4級
評分重點：詞彙運用、句式多樣、文句通順

具體標準：
- 4級(16-20分): 詞彙豐富，句式多樣，文句流暢，修辭恰當
- 3級(12-15分): 詞彙尚豐富，句式有變化，文句通順
- 2級(6-11分): 詞彙一般，句式單一，文句有語病
- 1級(0-5分): 詞彙貧乏，句式單調，文句不通

### 文類與格式（10分）- 等級制 1-4級
評分重點：符合文類要求、格式正確

具體標準：
- 4級(8-10分): 完全符合文類要求，格式正確
- 3級(6-7分): 大致符合文類要求，格式大致正確
- 2級(3-5分): 尚算符合文類要求，格式有錯誤
- 1級(0-2分): 不符合文類要求，格式錯誤嚴重

## 輸出格式
請以JSON格式返回：
{
  "grading": {
    "content": 3,
    "feeling": 3,
    "structure": 3,
    "language": 3,
    "format": 3
  },
  "overallComment": "簡潔易明的總評",
  "contentFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "feelingFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "structureFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "languageFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "formatFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "enhancedText": "增潤後的文章",
  "enhancementNotes": ["修改說明1"],
  "modelEssay": "示範文章"
}`;
}

function buildPracticalGradingPrompt() {
  return `你是一位專業的香港中學中文科教師，正在批改HKDSE中文卷二甲部實用寫作。

## 評分準則（必須嚴格遵守）

### 內容（30分）- 分為「資訊」和「拓展」兩部分

#### 資訊（最高6分）
評分重點：是否準確提取和運用題目提供的資訊

具體標準：
- 5-6分: 能準確提取所有重要資訊，並適當運用
- 3-4分: 能提取大部分資訊，運用尚算適當
- 1-2分: 資訊提取不完整，運用不當
- 0分: 未能提取或運用資訊

#### 拓展（最高16分）
評分重點：內容發展的充分性和說服力

具體標準：
- 13-16分: 內容發展極為充分，論點清晰有力，說服力強
- 9-12分: 內容發展充分，論點清晰，有說服力
- 5-8分: 內容發展一般，論點尚算清晰
- 1-4分: 內容發展不足，論點模糊
- 0分: 幾乎沒有內容發展

### 行文組織（20分）- 分為「語氣」和「組織」兩部分

#### 語氣（最高10分）
評分重點：語氣是否切合寫作情境、對象和目的

具體標準：
- 8-10分: 語氣極為切合，能因應情境、對象和目的調整
- 5-7分: 語氣大致切合，偶有偏差
- 2-4分: 語氣尚算切合，但有明顯偏差
- 0-1分: 語氣不切合，嚴重影響溝通效果

#### 組織（最高10分）
評分重點：結構是否清晰、條理分明

具體標準：
- 8-10分: 結構極為清晰，條理分明，過渡自然
- 5-7分: 結構清晰，條理尚算分明
- 2-4分: 結構一般，條理不夠清晰
- 0-1分: 結構散亂，缺乏條理

### 格式要求
不同文體有不同的格式要求，請根據具體文體評估格式是否正確。

## 輸出格式
請以JSON格式返回：
{
  "grading": {
    "info": 5,
    "development": 12,
    "tone": 8,
    "organization": 8
  },
  "overallComment": "簡潔易明的總評",
  "infoFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "developmentFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "toneFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "organizationFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善建議1"]
  },
  "formatIssues": ["格式問題1", "格式問題2"],
  "enhancedText": "增潤後的文章",
  "enhancementNotes": ["修改說明1"],
  "modelEssay": "示範文章"
}`;
}

function buildUserPrompt(essayText, question, customCriteria, gradingMode = 'secondary') {
  return `請批改以下學生作文：

## 題目
${question || '（未提供題目）'}

## 學生作文
${essayText}

${customCriteria ? `## 自定義批改準則\n${customCriteria}\n\n請將上述自定義準則與系統評分準則結合使用。` : ''}

請嚴格按照評分準則進行批改，給出具體、有建設性的評語和建議，並以JSON格式返回結果。`;
}

function parseGradingResult(content, essayText, gradingMode = 'secondary') {
  const result = safeJSONParse(content, 'grade');

  if (gradingMode === 'primary') {
    const grading = {
      content: Math.max(1, Math.min(4, Math.round(result.grading?.content || 3))) as 1|2|3|4,
      feeling: Math.max(1, Math.min(4, Math.round(result.grading?.feeling || 3))) as 1|2|3|4,
      structure: Math.max(1, Math.min(4, Math.round(result.grading?.structure || 3))) as 1|2|3|4,
      language: Math.max(1, Math.min(4, Math.round(result.grading?.language || 3))) as 1|2|3|4,
      format: Math.max(1, Math.min(4, Math.round(result.grading?.format || 3))) as 1|2|3|4,
    };

    const scoreTable = {
      content: [0, 9, 18, 24, 30],
      feeling: [0, 6, 12, 16, 20],
      structure: [0, 6, 12, 16, 20],
      language: [0, 6, 12, 16, 20],
      format: [0, 3, 6, 8, 10],
    };

    const totalScore = scoreTable.content[grading.content] + 
                      scoreTable.feeling[grading.feeling] +
                      scoreTable.structure[grading.structure] +
                      scoreTable.language[grading.language] +
                      scoreTable.format[grading.format];

    const buildFeedback = (fb) => ({
      strengths: Array.isArray(fb?.strengths) ? fb.strengths : [],
      improvements: Array.isArray(fb?.improvements) ? fb.improvements : []
    });

    return {
      grading,
      totalScore,
      gradeLevel: getPrimaryGradeLabel(totalScore),
      overallComment: result.overallComment || '',
      contentFeedback: buildFeedback(result.contentFeedback),
      feelingFeedback: buildFeedback(result.feelingFeedback),
      structureFeedback: buildFeedback(result.structureFeedback),
      languageFeedback: buildFeedback(result.languageFeedback),
      formatFeedback: buildFeedback(result.formatFeedback),
      enhancedText: result.enhancedText || essayText,
      enhancementNotes: Array.isArray(result.enhancementNotes) ? result.enhancementNotes : [],
      modelEssay: result.modelEssay || ''
    };
  } else if (gradingMode === 'practical') {
    const grading = {
      info: Math.max(0, Math.min(6, Math.round(result.grading?.info || 4))),
      development: Math.max(0, Math.min(16, Math.round(result.grading?.development || 10))),
      tone: Math.max(0, Math.min(10, Math.round(result.grading?.tone || 6))),
      organization: Math.max(0, Math.min(10, Math.round(result.grading?.organization || 6))),
    };

    const contentScore = (grading.info + grading.development);
    const organizationScore = grading.tone + grading.organization;
    const totalScore = contentScore + organizationScore;

    const buildFeedback = (fb) => ({
      strengths: Array.isArray(fb?.strengths) ? fb.strengths : [],
      improvements: Array.isArray(fb?.improvements) ? fb.improvements : []
    });

    return {
      grading,
      contentScore,
      organizationScore,
      totalScore,
      overallComment: result.overallComment || '',
      infoFeedback: buildFeedback(result.infoFeedback),
      developmentFeedback: buildFeedback(result.developmentFeedback),
      toneFeedback: buildFeedback(result.toneFeedback),
      organizationFeedback: buildFeedback(result.organizationFeedback),
      formatIssues: Array.isArray(result.formatIssues) ? result.formatIssues : [],
      enhancedText: result.enhancedText || essayText,
      enhancementNotes: Array.isArray(result.enhancementNotes) ? result.enhancementNotes : [],
      modelEssay: result.modelEssay || ''
    };
  } else {
    // secondary
    const grading = {
      content: Math.max(1, Math.min(10, Math.round(result.grading?.content || 6))),
      expression: Math.max(1, Math.min(10, Math.round(result.grading?.expression || 6))),
      structure: Math.max(1, Math.min(10, Math.round(result.grading?.structure || 6))),
      punctuation: Math.max(5, Math.min(10, Math.round(result.grading?.punctuation || 7)))
    };

    const totalScore = (grading.content * 4) + (grading.expression * 3) + 
                      (grading.structure * 2) + grading.punctuation;

    const buildFeedback = (fb) => ({
      strengths: Array.isArray(fb?.strengths) ? fb.strengths : [],
      improvements: Array.isArray(fb?.improvements) ? fb.improvements : []
    });

    const getGradeLabel = (score) => {
      if (score >= 90) return '上上';
      if (score >= 85) return '上中';
      if (score >= 80) return '上下';
      if (score >= 70) return '中上';
      if (score >= 60) return '中中';
      if (score >= 50) return '中下';
      if (score >= 40) return '下上';
      if (score >= 30) return '下中';
      if (score >= 10) return '下下';
      return '極差';
    };

    return {
      grading,
      totalScore,
      gradeLabel: getGradeLabel(totalScore),
      overallComment: result.overallComment || '',
      contentFeedback: buildFeedback(result.contentFeedback),
      expressionFeedback: buildFeedback(result.expressionFeedback),
      structureFeedback: buildFeedback(result.structureFeedback),
      punctuationFeedback: buildFeedback(result.punctuationFeedback),
      enhancedText: result.enhancedText || essayText,
      enhancementNotes: Array.isArray(result.enhancementNotes) ? result.enhancementNotes : [],
      modelEssay: result.modelEssay || ''
    };
  }
}

function getPrimaryGradeLabel(totalScore) {
  if (totalScore >= 85) return '優異';
  if (totalScore >= 70) return '良好';
  if (totalScore >= 50) return '一般';
  return '有待改善';
}

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`========================================`);
  console.log(`中文科批改系統 API 服務器`);
  console.log(`運行於端口: ${PORT}`);
  console.log(`========================================`);
});

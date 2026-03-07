const express = require('express');
const cors = require('cors');
const multer = require('multer');
require('dotenv').config();

const app = express();
const upload = multer({ storage: multer.memoryStorage() });

// 啟用 CORS - 允許前端訪問
app.use(cors({
  origin: '*', // 允許所有來源（開發時使用）
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
        return res.status(400).json({ 
          success: false, 
          message: '自定義 API 需要提供 API 基礎 URL' 
        });
      }
      const result = await testCustomConnection(apiKey, baseURL, model);
      return res.json(result);
    } else if (apiType === 'openai') {
      const result = await testOpenAIConnection(apiKey, model);
      return res.json(result);
    } else {
      return res.status(400).json({ 
        success: false, 
        message: '未知的 API 類型: ' + apiType 
      });
    }
  } catch (error) {
    console.error('Test API error:', error);
    res.status(500).json({ 
      success: false, 
      message: error.message || '測試失敗'
    });
  }
});

// OCR 提取文字
app.post('/api/extract', upload.single('file'), async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, fileType, fileData, text } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    if (!fileData && !text) {
      return res.status(400).json({ success: false, message: '沒有提供文件或文字' });
    }

    let result;
    if (apiType === 'gemini') {
      result = await extractWithGemini(apiKey, model, fileData, text, fileType);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ 
          success: false, 
          message: '自定義 API 需要提供 API 基礎 URL' 
        });
      }
      result = await extractWithCustom(apiKey, baseURL, model, fileData, text, fileType);
    } else if (apiType === 'openai') {
      result = await extractWithOpenAI(apiKey, model, fileData, text, fileType);
    } else {
      return res.status(400).json({ 
        success: false, 
        message: '未知的 API 類型: ' + apiType 
      });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Extract error:', error);
    res.status(500).json({ 
      success: false, 
      message: error.message || '提取失敗'
    });
  }
});

// 批改作文
app.post('/api/grade', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, essayText, question, customCriteria } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    if (!essayText) {
      return res.status(400).json({ success: false, message: '作文內容不能為空' });
    }

    let result;
    if (apiType === 'gemini') {
      result = await gradeWithGemini(apiKey, model, essayText, question, customCriteria);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ 
          success: false, 
          message: '自定義 API 需要提供 API 基礎 URL' 
        });
      }
      result = await gradeWithCustom(apiKey, baseURL, model, essayText, question, customCriteria);
    } else if (apiType === 'openai') {
      result = await gradeWithOpenAI(apiKey, model, essayText, question, customCriteria);
    } else {
      return res.status(400).json({ 
        success: false, 
        message: '未知的 API 類型: ' + apiType 
      });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Grade error:', error);
    res.status(500).json({ 
      success: false, 
      message: error.message || '批改失敗'
    });
  }
});

// 全班分析
app.post('/api/analyze-class', async (req, res) => {
  try {
    const { apiKey, apiType, model, baseURL, reports, question } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ success: false, message: 'API 密鑰不能為空' });
    }

    let result;
    if (apiType === 'gemini') {
      result = await analyzeClassWithGemini(apiKey, model, reports, question);
    } else if (apiType === 'custom') {
      if (!baseURL) {
        return res.status(400).json({ 
          success: false, 
          message: '自定義 API 需要提供 API 基礎 URL' 
        });
      }
      result = await analyzeClassWithCustom(apiKey, baseURL, model, reports, question);
    } else if (apiType === 'openai') {
      result = await analyzeClassWithOpenAI(apiKey, model, reports, question);
    } else {
      return res.status(400).json({ 
        success: false, 
        message: '未知的 API 類型: ' + apiType 
      });
    }

    res.json({ success: true, ...result });
  } catch (error) {
    console.error('Analyze class error:', error);
    res.status(500).json({ 
      success: false, 
      message: error.message || '分析失敗'
    });
  }
});

// ============ Gemini API 函數 ============

// 驗證和修正 Gemini 模型名稱
function normalizeGeminiModelName(modelName) {
  if (!modelName) return 'gemini-2.0-flash';
  
  // 移除 models/ 前綴（如果存在）
  let name = modelName.startsWith('models/') ? modelName.substring(7) : modelName;
  
  // 常見的模型名稱映射
  const modelMap = {
    'gemini-1.5-flash': 'gemini-1.5-flash-latest',
    'gemini-1.5-pro': 'gemini-1.5-pro-latest',
    'gemini-1.0-pro': 'gemini-1.0-pro-latest',
    'gemini-pro': 'gemini-1.0-pro-latest',
  };
  
  // 如果名稱在映射中，使用映射後的名稱
  if (modelMap[name]) {
    return modelMap[name];
  }
  
  // 如果名稱已經包含 -latest 或版本號，直接返回
  if (name.includes('-latest') || /-\d{3}$/.test(name)) {
    return name;
  }
  
  // 否則返回原始名稱
  return name;
}

async function testGeminiConnection(apiKey, modelName = 'gemini-2.0-flash') {
  // 修正模型名稱
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
    const isModelAvailable = availableModels.some(m => m === model || m.endsWith(modelName));

    return {
      success: true,
      message: isModelAvailable 
        ? `Gemini API 連接成功！模型 "${modelName}" 可用`
        : `API 連接成功！但模型 "${modelName}" 可能不可用。可用模型: ${availableModels.slice(0, 3).join(', ')}`,
      model: modelName
    };
  } catch (error) {
    return handleGeminiError(error, modelName);
  }
}

async function extractWithGemini(apiKey, modelName, fileData, text, fileType) {
  // 修正模型名稱
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  const prompt = `你是一個專業的OCR文字識別助手。請從圖片或文檔中提取學生的作文文字。

提取要求：
1. 保持原文的段落格式
2. 識別學生姓名和學號（通常在文章開頭或標題處）
3. 只提取作文正文，不要包含題目（除非題目是文章的一部分）
4. 保持所有標點符號
5. 不要修改任何文字，包括錯別字

請以JSON格式返回：
{
  "text": "提取的作文全文",
  "name": "學生姓名（如無則留空）",
  "studentId": "學生學號（如無則留空）"
}`;

  let requestBody;

  if (fileData && fileType && fileType.startsWith('image/')) {
    // 圖片處理
    requestBody = {
      contents: [{
        parts: [
          { text: prompt },
          {
            inline_data: {
              mime_type: fileType,
              data: fileData
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 4000,
        responseMimeType: 'application/json'
      }
    };
  } else {
    // 文字處理
    const content = text || '';
    requestBody = {
      contents: [{
        parts: [{
          text: prompt + '\n\n請提取以下文本中的學生作文內容：\n\n' + content
        }]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 4000,
        responseMimeType: 'application/json'
      }
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
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    throw new Error('Gemini API 返回內容為空');
  }

  const jsonContent = content.replace(/```json\n?|\n?```/g, '').trim();
  const result = JSON.parse(jsonContent);
  
  return {
    text: result.text || '',
    name: result.name || '',
    studentId: result.studentId || ''
  };
}

async function gradeWithGemini(apiKey, modelName, essayText, question, customCriteria) {
  // 修正模型名稱
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  const systemPrompt = buildGradingPrompt();
  const userPrompt = buildUserPrompt(essayText, question, customCriteria);

  const requestBody = {
    contents: [{
      parts: [{ text: systemPrompt + '\n\n' + userPrompt }]
    }],
    generationConfig: {
      temperature: 0.5,
      maxOutputTokens: 8000,
      responseMimeType: 'application/json'
    }
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
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    throw new Error('Gemini API 返回內容為空');
  }

  const jsonContent = content.replace(/```json\n?|\n?```/g, '').trim();
  return parseGradingResult(jsonContent, essayText);
}

async function analyzeClassWithGemini(apiKey, modelName, reports, question) {
  // 修正模型名稱
  const normalizedModelName = normalizeGeminiModelName(modelName);
  const model = `models/${normalizedModelName}`;
  
  const stats = {
    totalStudents: reports.length,
    averageScore: reports.reduce((sum, r) => sum + r.totalScore, 0) / reports.length,
    contentAvg: reports.reduce((sum, r) => sum + r.grading.content, 0) / reports.length,
    expressionAvg: reports.reduce((sum, r) => sum + r.grading.expression, 0) / reports.length,
    structureAvg: reports.reduce((sum, r) => sum + r.grading.structure, 0) / reports.length,
  };

  const systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的寫作表現。

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

  const userPrompt = `請分析以下全班寫作數據：

## 題目
${question}

## 統計數據
- 總人數: ${stats.totalStudents}
- 平均分: ${stats.averageScore.toFixed(1)}
- 內容平均分: ${stats.contentAvg.toFixed(1)}
- 表達平均分: ${stats.expressionAvg.toFixed(1)}
- 結構平均分: ${stats.structureAvg.toFixed(1)}

## 學生分數詳情
${reports.map(r => `- ${r.studentWork?.name || '未命名'}: 總分${r.totalScore} (內容${r.grading.content}, 表達${r.grading.expression}, 結構${r.grading.structure}, 標點${r.grading.punctuation})`).join('\n')}

請以專業教師的角度進行分析，並以JSON格式返回結果。`;

  const requestBody = {
    contents: [{
      parts: [{ text: systemPrompt + '\n\n' + userPrompt }]
    }],
    generationConfig: {
      temperature: 0.5,
      maxOutputTokens: 8000,
      responseMimeType: 'application/json'
    }
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
  const content = data.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!content) {
    throw new Error('Gemini API 返回內容為空');
  }

  const jsonContent = content.replace(/```json\n?|\n?```/g, '').trim();
  return JSON.parse(jsonContent);
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
        : `API 連接成功！但模型 "${modelName}" 可能不可用。可用模型: ${availableModels.slice(0, 3).join(', ')}`,
      model: modelName
    };
  } catch (error) {
    return handleOpenAIError(error, modelName);
  }
}

async function extractWithOpenAI(apiKey, modelName, fileData, text, fileType) {
  const model = modelName || 'gpt-4o';
  
  const systemPrompt = `你是一個專業的OCR文字識別助手。請從圖片或文檔中提取學生的作文文字。

提取要求：
1. 保持原文的段落格式
2. 識別學生姓名和學號（通常在文章開頭或標題處）
3. 只提取作文正文，不要包含題目（除非題目是文章的一部分）
4. 保持所有標點符號
5. 不要修改任何文字，包括錯別字

請以JSON格式返回：
{
  "text": "提取的作文全文",
  "name": "學生姓名（如無則留空）",
  "studentId": "學生學號（如無則留空）"
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
  return {
    text: result.text || '',
    name: result.name || '',
    studentId: result.studentId || ''
  };
}

async function gradeWithOpenAI(apiKey, modelName, essayText, question, customCriteria) {
  const model = modelName || 'gpt-4o';
  
  const systemPrompt = buildGradingPrompt();
  const userPrompt = buildUserPrompt(essayText, question, customCriteria);

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
      max_tokens: 8000,
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

  return parseGradingResult(content, essayText);
}

async function analyzeClassWithOpenAI(apiKey, modelName, reports, question) {
  const model = modelName || 'gpt-4o';
  
  const stats = {
    totalStudents: reports.length,
    averageScore: reports.reduce((sum, r) => sum + r.totalScore, 0) / reports.length,
    contentAvg: reports.reduce((sum, r) => sum + r.grading.content, 0) / reports.length,
    expressionAvg: reports.reduce((sum, r) => sum + r.grading.expression, 0) / reports.length,
    structureAvg: reports.reduce((sum, r) => sum + r.grading.structure, 0) / reports.length,
  };

  const systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的寫作表現。

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

  const userPrompt = `請分析以下全班寫作數據：

## 題目
${question}

## 統計數據
- 總人數: ${stats.totalStudents}
- 平均分: ${stats.averageScore.toFixed(1)}
- 內容平均分: ${stats.contentAvg.toFixed(1)}
- 表達平均分: ${stats.expressionAvg.toFixed(1)}
- 結構平均分: ${stats.structureAvg.toFixed(1)}

## 學生分數詳情
${reports.map(r => `- ${r.studentWork?.name || '未命名'}: 總分${r.totalScore} (內容${r.grading.content}, 表達${r.grading.expression}, 結構${r.grading.structure}, 標點${r.grading.punctuation})`).join('\n')}

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
      max_tokens: 8000,
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

// ============ 自定義 API 函數（支持任何 OpenAI 兼容 API） ============

async function testCustomConnection(apiKey, baseURL, modelName) {
  try {
    // 確保 baseURL 以 /v1 結尾
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
        : `API 連接成功！但模型 "${modelName}" 可能不可用。可用模型: ${availableModels.slice(0, 3).join(', ')}`,
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
  "text": "提取的作文全文",
  "name": "學生姓名（如無則留空）",
  "studentId": "學生學號（如無則留空）"
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
  return {
    text: result.text || '',
    name: result.name || '',
    studentId: result.studentId || ''
  };
}

async function gradeWithCustom(apiKey, baseURL, modelName, essayText, question, customCriteria) {
  const model = modelName || 'gpt-4o';
  const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
  
  const systemPrompt = buildGradingPrompt();
  const userPrompt = buildUserPrompt(essayText, question, customCriteria);

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
      max_tokens: 8000,
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

  return parseGradingResult(content, essayText);
}

async function analyzeClassWithCustom(apiKey, baseURL, modelName, reports, question) {
  const model = modelName || 'gpt-4o';
  const normalizedURL = baseURL.endsWith('/v1') ? baseURL : `${baseURL}/v1`;
  
  const stats = {
    totalStudents: reports.length,
    averageScore: reports.reduce((sum, r) => sum + r.totalScore, 0) / reports.length,
    contentAvg: reports.reduce((sum, r) => sum + r.grading.content, 0) / reports.length,
    expressionAvg: reports.reduce((sum, r) => sum + r.grading.expression, 0) / reports.length,
    structureAvg: reports.reduce((sum, r) => sum + r.grading.structure, 0) / reports.length,
  };

  const systemPrompt = `你是一位專業的香港中學中文科教師，正在分析全班學生的寫作表現。

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

  const userPrompt = `請分析以下全班寫作數據：

## 題目
${question}

## 統計數據
- 總人數: ${stats.totalStudents}
- 平均分: ${stats.averageScore.toFixed(1)}
- 內容平均分: ${stats.contentAvg.toFixed(1)}
- 表達平均分: ${stats.expressionAvg.toFixed(1)}
- 結構平均分: ${stats.structureAvg.toFixed(1)}

## 學生分數詳情
${reports.map(r => `- ${r.studentWork?.name || '未命名'}: 總分${r.totalScore} (內容${r.grading.content}, 表達${r.grading.expression}, 結構${r.grading.structure}, 標點${r.grading.punctuation})`).join('\n')}

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
      max_tokens: 8000,
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

// ============ 輔助函數 ============

function handleGeminiError(error, modelName) {
  const message = error.message || '未知錯誤';
  
  if (message.includes('quota') || message.includes('exceeded') || message.includes('429')) {
    return { 
      success: false, 
      message: 'Gemini API 配額已用完，請檢查您的 Google AI Studio 配額或稍後再試' 
    };
  }
  
  if (message.includes('not found') || message.includes('model')) {
    return { 
      success: false, 
      message: `模型 "${modelName}" 不存在或無權訪問，請確認模型名稱正確` 
    };
  }
  
  if (message.includes('401') || message.includes('Unauthorized') || message.includes('API key not valid')) {
    return { 
      success: false, 
      message: 'API 密鑰無效，請檢查您的 Gemini API 密鑰是否正確' 
    };
  }
  
  if (message.includes('fetch') || message.includes('network') || message.includes('ECONNREFUSED')) {
    return { 
      success: false, 
      message: '無法連接到 Gemini 服務器，請檢查網絡連接' 
    };
  }
  
  return { success: false, message: `連接失敗: ${message}` };
}

function handleOpenAIError(error, modelName) {
  const message = error.message || '未知錯誤';
  
  if (message.includes('quota') || message.includes('exceeded') || message.includes('429')) {
    return { 
      success: false, 
      message: 'API 配額已用完，請檢查您的賬戶配額' 
    };
  }
  
  if (message.includes('not found') || message.includes('model')) {
    return { 
      success: false, 
      message: `模型 "${modelName}" 不存在或無權訪問` 
    };
  }
  
  if (message.includes('401') || message.includes('Unauthorized')) {
    return { 
      success: false, 
      message: 'API 密鑰無效，請檢查您的 API 密鑰' 
    };
  }
  
  return { success: false, message: `連接失敗: ${message}` };
}

function buildGradingPrompt() {
  return `你是一位專業的香港中學中文科教師，正在批改HKDSE中文卷二乙部命題寫作。

## 評分準則

### 內容（40分）- 品第制 1-10分
- 上上(10): 立意極豐富深刻，取材極恰當，闡述極飽滿
- 上中(9): 立意極豐富尚算深刻，取材極恰當，闡述飽滿
- 上下(8): 立意豐富尚算深刻，取材恰當，闡述飽滿
- 中上(7): 立意平穩具體，取材合理平穩，闡述合理
- 中中上(6): 立意一般恰當，取材一般，闡述一般
- 中中下(5): 立意尚具單薄，取材單薄，闡述浮淺
- 中下(4): 立意十分單薄，取材薄弱，闡述極浮淺
- 下上(3): 立意模糊不太相干，取材不扣連，闡述欠奉
- 下中(2): 立意十分模糊錯亂，取材混亂，闡述空泛
- 下下(1): 無立意或錯亂極多，取材闡述闕如

### 表達（30分）- 品第制 1-10分
- 上上(10): 用詞極精確豐富，文句極簡潔流暢，手法純熟靈活
- 上中(9): 用詞極精確豐富，文句極簡潔流暢，手法純熟
- 上下(8): 用詞精確豐富，文句簡潔流暢，手法靈活
- 中上(7): 用詞準確平穩，文句通順偶有瑕疵，手法尚算靈活
- 中中上(6): 用詞大致準確，文句大致通順有沙石，表達一般
- 中中下(5): 用詞尚算準確，文句尚算通順有明顯語病，表達浮淺
- 中下(4): 用詞粗疏，文句不通順失誤多，表達薄弱
- 下上(3): 用詞不準，文句欠通順，表達混亂
- 下中(2): 用詞句式嚴重錯誤，表達極混亂
- 下下(1): 文句無法達意

### 結構（20分）- 品第制 1-10分
- 上上(10): 結構極完整，詳略極得宜，鋪排主次有序
- 上中(9): 結構極完整，詳略得宜，鋪排有序
- 上下(8): 結構完整，詳略得宜，鋪排有序
- 中上(7): 結構大致完整，詳略大致合宜，過渡尚算自然
- 中中上(6): 結構尚具完整，詳略一般有輕微失衡
- 中中下(5): 結構尚具，詳略稍失衡，偶有散亂
- 中下(4): 尚具組織，詳略明顯失衡，鋪排失當
- 下上(3): 組織散亂，詳略嚴重失衡
- 下中(2): 毫無組織，鋪排極混亂
- 下下(1): 完全無結構

### 標點（10分）- 5-10分
- 10分: 標點符號運用正確無誤
- 9分: 偶有失誤
- 8分: 有少許失誤
- 7分: 有失誤
- 6分: 失誤較多
- 5分: 有明顯問題

## 重要規則
1. 內容與結構分數一般不應相差超過2級
2. 若內容離題（3分及以下），結構最高只能評至7分
3. 標點分數下限為5分

## 輸出格式
請以JSON格式返回，確保所有字段都存在：
{
  "grading": {
    "content": 6,
    "expression": 6,
    "structure": 6,
    "punctuation": 7
  },
  "overallComment": "簡潔易明的總評，2-3段",
  "contentFeedback": {
    "strengths": ["優點1", "優點2"],
    "improvements": ["改善1", "改善2"]
  },
  "expressionFeedback": {
    "strengths": ["優點1", "優點2"],
    "improvements": ["改善1", "改善2"]
  },
  "structureFeedback": {
    "strengths": ["優點1", "優點2"],
    "improvements": ["改善1", "改善2"]
  },
  "punctuationFeedback": {
    "strengths": ["優點1"],
    "improvements": ["改善1"]
  },
  "enhancedText": "增潤後的完整文章，避免AI堆砌感，要有人文情懷",
  "enhancementNotes": ["修改說明1", "修改說明2"],
  "modelEssay": "一篇符合HKDSE上上品標準的奪星文章示範"
}`;
}

function buildUserPrompt(essayText, question, customCriteria) {
  return `請批改以下學生作文：

## 題目
${question || '（未提供題目）'}

## 學生作文
${essayText}

${customCriteria ? `## 自定義批改準則\n${customCriteria}` : ''}

請根據HKDSE中文卷二乙部命題寫作評分準則進行批改，並以JSON格式返回結果。`;
}

function parseGradingResult(content, essayText) {
  const result = JSON.parse(content);

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

// 啟動服務器
const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`========================================`);
  console.log(`中文科批改系統 API 服務器`);
  console.log(`運行於端口: ${PORT}`);
  console.log(`========================================`);
});

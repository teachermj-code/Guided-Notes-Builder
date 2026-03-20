/***************************************
 * LESSON GUIDE BUILDER BY TEACHER MJ
 * Server-side Apps Script
 ***************************************/

const APP_TITLE = 'Lesson Guide Builder by Teacher MJ';
const DEFAULT_MODEL = 'models/gemini-flash-latest';
const DEFAULT_LAYOUT_MODE = 'STANDARD';
const DEFAULT_TEMPLATE = 'CONCEPT';
const DEFAULT_ITEM_COUNT = 10;

/***************************************
 * 1. WEB APP ENTRY
 ***************************************/
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/***************************************
 * 2. PUBLIC ACTIONS
 ***************************************/
function getLessonStyles_() {
  return [
    '.practice-number{font-size:12px;font-weight:700;color:#5b6b82;text-transform:uppercase;letter-spacing:.06em;margin-bottom:10px;}',
    '.guided-summary-box{background:var(--template-subtle-bg);border:1px solid var(--template-subtle-border);border-radius:12px;padding:10px 12px;}',
    '.guided-symbol-box{background:#ffffff;border:1px dashed var(--template-soft-border);border-radius:12px;padding:10px 12px;margin-top:10px;}',
    '.guided-focus-strip{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:14px;padding:14px 16px;margin-bottom:14px;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .guided-learn-card{background:#fff;border:1px solid var(--template-subtle-border);box-shadow:none;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .guided-learn-card .guided-summary-box{background:transparent;border:1px solid var(--template-subtle-border);padding:10px 12px;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .guided-learn-card .guided-symbol-box{background:#fff;border:1px dashed var(--template-subtle-border);padding:10px 12px;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .guided-learn-card .misconception-box{margin-top:10px;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .guided-example-card{background:#fff;border:1px solid var(--template-soft-border);border-left:5px solid var(--template-accent);box-shadow:none;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .practice-section{margin-top:8px;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .practice-section .section-title{font-size:22px;letter-spacing:.06em;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .practice-section .section-note{color:var(--template-accent-dark);font-weight:600;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .practice-card{background:#fff;border:1px solid var(--template-soft-border);border-left:5px solid var(--template-accent);border-radius:16px;box-shadow:0 8px 18px rgba(15,39,71,.08);}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .practice-card .practice-number{color:var(--template-accent-dark);letter-spacing:.08em;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .practice-card .practice-question{font-size:15px;line-height:1.75;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"] .practice-card .teacher-answer-box{background:var(--template-soft-bg);border:1px dashed var(--template-soft-border);}',
    '.guided-example-title{font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--template-accent-dark);margin-bottom:12px;}',
    '.lesson-meta{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-top:18px;}',
    '.lesson-sheet{--template-accent:#0052cc;--template-accent-dark:#0f2747;--template-soft-bg:#f5f9ff;--template-soft-border:#d9e6ff;--template-subtle-bg:#fafcff;--template-subtle-border:#e5ebf5;max-width:900px;margin:0 auto;padding:24px;}',
    '.lesson-sheet.layout-compact{padding:18px;}',
    '.lesson-sheet.layout-spacious{padding:30px;}',
    '.lesson-sheet[data-template="CONCEPT"]{--template-accent:#0052cc;--template-accent-dark:#0f2747;--template-soft-bg:#f5f9ff;--template-soft-border:#d9e6ff;--template-subtle-bg:#fafcff;--template-subtle-border:#e5ebf5;}',
    '.lesson-sheet[data-template="GUIDED_PRACTICE"]{--template-accent:#0f766e;--template-accent-dark:#134e4a;--template-soft-bg:#f0fdfa;--template-soft-border:#99f6e4;--template-subtle-bg:#f7fffd;--template-subtle-border:#c7f3eb;}',
    '.lesson-sheet[data-template="REVIEW"]{--template-accent:#7c3aed;--template-accent-dark:#4c1d95;--template-soft-bg:#f5f3ff;--template-soft-border:#ddd6fe;--template-subtle-bg:#faf8ff;--template-subtle-border:#ece7ff;}',
    '.lesson-sheet[data-template="QUIZ"]{--template-accent:#c2410c;--template-accent-dark:#7c2d12;--template-soft-bg:#fff7ed;--template-soft-border:#fed7aa;--template-subtle-bg:#fffaf5;--template-subtle-border:#ffedd5;}',
    '.lesson-sheet[data-template="REMEDIATION"]{--template-accent:#be185d;--template-accent-dark:#831843;--template-soft-bg:#fff1f2;--template-soft-border:#fecdd3;--template-subtle-bg:#fff8f8;--template-subtle-border:#ffe4e6;}',
    '.lesson-sheet[data-template="ENRICHMENT"]{--template-accent:#047857;--template-accent-dark:#064e3b;--template-soft-bg:#ecfdf5;--template-soft-border:#a7f3d0;--template-subtle-bg:#f5fffa;--template-subtle-border:#d1fae5;}',
    '.lesson-header{background:#ffffff;border:1px solid #dfe3eb;border-top:5px solid var(--template-accent);border-radius:18px;padding:18px;margin-bottom:14px;}',
    '.lesson-sheet.layout-compact .lesson-header{padding:14px;margin-bottom:10px;}',
    '.lesson-sheet.layout-spacious .lesson-header{padding:22px;margin-bottom:18px;}',
    '.doc-banner{margin-bottom:18px;}',
    '.school-name{font-size:15px;font-weight:700;color:var(--template-accent-dark);margin-bottom:4px;}',
    '.custom-header{font-size:12px;color:#5b6b82;margin-bottom:8px;}',
    '.eyebrow{font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--template-accent);margin-bottom:8px;}',
    '.lesson-title{margin:0;font-size:28px;line-height:1.2;color:var(--template-accent-dark);}',
    '.lesson-meta>div{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:12px;padding:12px;}',
    '.meta-label{display:block;font-size:11px;font-weight:700;text-transform:uppercase;color:#5b6b82;margin-bottom:4px;}',
    '.meta-value{display:block;font-size:14px;font-weight:600;color:#1c2e45;}',
    '.section-title{font-size:18px;margin:24px 0 6px;color:var(--template-accent);}',
    '.section-note{font-size:13px;color:#5b6b82;margin-bottom:14px;line-height:1.6;}',
    '.card{background:#ffffff;border:1px solid #dfe3eb;border-radius:16px;padding:18px;margin-bottom:14px;}',
    '.lesson-sheet.layout-compact .card{padding:14px;margin-bottom:10px;}',
    '.lesson-sheet.layout-spacious .card{padding:22px;margin-bottom:18px;}',
    '.concept-card,.practice-card{break-inside:avoid;page-break-inside:avoid;}',
    '.concept-title{margin:0 0 10px;font-size:16px;color:#102a43;line-height:1.35;}',
    '.concept-summary{margin:0 0 12px;font-size:14px;line-height:1.7;}',
    '.formula-box{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:12px;padding:12px;margin-bottom:12px;}',
    '.formula-label{font-size:11px;font-weight:700;text-transform:uppercase;color:#5b6b82;margin-bottom:6px;}',
    '.formula-value{font-size:15px;line-height:1.7;}',
    '.sub-card{background:var(--template-subtle-bg);border:1px solid var(--template-subtle-border);border-radius:12px;padding:12px;margin-bottom:12px;font-size:13px;line-height:1.6;}',
    '.sub-card-label{font-size:11px;font-weight:700;text-transform:uppercase;color:#5b6b82;margin-bottom:6px;}',
    '.example-box{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:12px;padding:12px;font-size:13px;line-height:1.7;margin-bottom:10px;}',
    '.misconception-box{background:#fff8e6;border:1px solid #f2d27c;border-radius:12px;padding:12px;font-size:13px;line-height:1.6;}',
    '.criteria-list,.teacher-note-list{margin:0;padding-left:20px;line-height:1.7;}',
    '.vocab-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;}',
    '.vocab-item{background:var(--template-subtle-bg);border:1px solid var(--template-subtle-border);border-radius:12px;padding:12px;}',
    '.vocab-term{font-weight:700;color:#102a43;margin-bottom:6px;}',
    '.vocab-definition{font-size:13px;line-height:1.6;}',
    '.compact-vocab-list{margin:0;padding-left:20px;line-height:1.7;}',
    '.compact-vocab-list li{margin-bottom:8px;}',
    '.compact-vocab-term{font-weight:700;color:#102a43;}',
    '.review-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;}',
    '.review-focus-strip{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:14px;padding:14px 16px;margin-bottom:14px;}',
    '.review-focus-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--template-accent-dark);margin-bottom:6px;}',
    '.review-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px;}',
    '.remediation-strip,.enrichment-strip{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:14px;padding:14px 16px;margin-bottom:14px;}',
    '.remediation-strip-title,.enrichment-strip-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.06em;color:var(--template-accent-dark);margin-bottom:6px;}',
    '.remediation-grid,.enrichment-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;}',
    '.remediation-card,.enrichment-card{background:#fff;border:1px solid var(--template-subtle-border);border-radius:14px;padding:14px;box-shadow:none;break-inside:avoid;page-break-inside:avoid;}',
    '.remediation-card .concept-summary,.enrichment-card .concept-summary{font-size:13px;line-height:1.6;margin-bottom:8px;}',
    '.remediation-help-box,.enrichment-note-box{background:#fff;border:1px dashed var(--template-soft-border);border-radius:10px;padding:10px 12px;margin-top:8px;}',
    '.remediation-model-card{background:#fff;border:1px solid var(--template-soft-border);border-left:5px solid var(--template-accent);border-radius:14px;padding:16px;margin-bottom:12px;box-shadow:none;}',
    '.remediation-model-title,.enrichment-card-title{font-size:12px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:var(--template-accent-dark);margin-bottom:8px;}',
    '.lesson-sheet[data-template="REMEDIATION"] .practice-card{background:#fff;border:1px solid var(--template-soft-border);box-shadow:none;}',
    '.lesson-sheet[data-template="REMEDIATION"] .practice-card .teacher-answer-box{background:var(--template-soft-bg);border:1px dashed var(--template-soft-border);}',
    '.lesson-sheet[data-template="ENRICHMENT"] .practice-card{background:#fff;border:1px solid var(--template-soft-border);border-left:5px solid var(--template-accent);box-shadow:none;}',
    '.lesson-sheet[data-template="ENRICHMENT"] .challenge-section .practice-card{border-left-color:var(--template-accent);}',
    '.lesson-sheet[data-template="ENRICHMENT"] .transfer-section .practice-card{border-left-color:var(--template-accent-dark);}',
    '.lesson-sheet[data-template="ENRICHMENT"] .practice-card .teacher-answer-box{background:var(--template-soft-bg);border:1px dashed var(--template-soft-border);}',
    '.lesson-sheet[data-template="ENRICHMENT"] .practice-section .section-title{letter-spacing:.05em;}',
    '.review-mini-card{background:#fff;border:1px solid var(--template-subtle-border);border-radius:14px;padding:14px;box-shadow:none;}',
    '.review-mini-title{font-size:12px;font-weight:700;letter-spacing:.05em;text-transform:uppercase;color:var(--template-accent-dark);margin-bottom:8px;}',
    '.review-rule-box{background:var(--template-subtle-bg);border:1px solid var(--template-subtle-border);border-radius:10px;padding:10px 12px;margin:8px 0 10px;}',
    '.review-reminder-box{background:#fff;border:1px dashed var(--template-soft-border);border-radius:10px;padding:10px 12px;margin-top:8px;}',
    '.lesson-sheet[data-template="REVIEW"] .compact-vocab-list{columns:2;column-gap:24px;padding-left:18px;}',
    '.lesson-sheet[data-template="REVIEW"] .compact-vocab-list li{break-inside:avoid;margin-bottom:8px;}',
    '.lesson-sheet[data-template="REVIEW"] .quick-review-card{background:#fff;border:1px solid var(--template-subtle-border);padding:14px;box-shadow:none;}',
    '.lesson-sheet[data-template="REVIEW"] .quick-review-card .concept-summary{font-size:13px;line-height:1.6;margin-bottom:8px;}',
    '.lesson-sheet[data-template="REVIEW"] .quick-review-card .misconception-box{margin-top:8px;padding:10px 12px;}',
    '.lesson-sheet[data-template="REVIEW"] .practice-section .section-title{letter-spacing:.04em;}',
    '.lesson-sheet[data-template="REVIEW"] .practice-card{background:#fff;border:1px solid var(--template-soft-border);box-shadow:none;}',
    '.lesson-sheet[data-template="REVIEW"] .practice-card .practice-question{font-size:14px;line-height:1.65;}',
    '.guided-practice-card,.quick-review-card{break-inside:avoid;page-break-inside:avoid;}',
    '.mini-formula-box{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:12px;padding:10px 12px;margin:10px 0 12px;}',
    '.mini-label{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#5b6b82;margin-bottom:4px;}',
    '.compact-note{font-size:13px;line-height:1.6;}',
    '.directions-box{background:var(--template-soft-bg);border:1px solid var(--template-soft-border);border-radius:14px;padding:14px 16px;line-height:1.7;font-size:13px;margin-bottom:14px;}',
    '.directions-list{margin:8px 0 0;padding-left:20px;}',
    '.quiz-info-card{background:#fff;border:1px solid var(--template-soft-border);border-radius:14px;padding:14px 16px;margin-bottom:14px;}',
    '.quiz-meta-lines{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;}',
    '.quiz-line{display:flex;align-items:flex-end;gap:8px;min-height:28px;}',
    '.quiz-line-label{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em;color:#5b6b82;white-space:nowrap;}',
    '.quiz-line-fill{flex:1;border-bottom:1.5px solid #7a869a;height:20px;}',
    '.quiz-focus-card{background:var(--template-subtle-bg);border:1px solid var(--template-subtle-border);border-radius:14px;padding:14px 16px;margin-bottom:14px;}',
    '.quiz-focus-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:var(--template-accent-dark);margin-bottom:6px;}',
    '.quiz-focus-text{font-size:13px;line-height:1.6;}',
    '.quiz-focus-block{margin-top:12px;}',
    '.lesson-sheet[data-template="QUIZ"] .practice-section .section-note{margin-bottom:10px;}',
    '.lesson-sheet[data-template="QUIZ"] .practice-card{background:#fff;border:1px solid #ead9cb;border-radius:12px;padding:16px;margin-bottom:10px;box-shadow:none;}',
    '.lesson-sheet[data-template="QUIZ"] .practice-number{font-size:11px;letter-spacing:.08em;color:var(--template-accent-dark);margin-bottom:10px;}',
    '.lesson-sheet[data-template="QUIZ"] .practice-question{font-size:15px;line-height:1.75;margin-bottom:14px;}',
    '.lesson-sheet[data-template="QUIZ"] .mc-choice{background:#fff;border-color:#e7d7c7;}',
    '.lesson-sheet[data-template="QUIZ"] .option-preview{background:#fffaf5;border-color:#edd5c4;}',
    '.lesson-sheet[data-template="QUIZ"] .teacher-answer-box{background:#fffaf5;border:1px solid #edd5c4;}',
    '.practice-number{font-size:12px;font-weight:700;color:#5b6b82;text-transform:uppercase;margin-bottom:8px;}',
    '.practice-question{font-size:14px;line-height:1.7;margin-bottom:12px;}',
    '.practice-controls{display:flex;flex-wrap:wrap;gap:10px;align-items:center;}',
    '.practice-controls-block{display:block;width:100%;}',
    '.mc-choice-list{display:flex;flex-direction:column;gap:10px;width:100%;}',
    '.mc-choice{display:flex;align-items:flex-start;gap:10px;padding:10px 12px;border:1px solid var(--template-subtle-border);border-radius:12px;background:var(--template-subtle-bg);}',
    '.mc-letter{font-weight:700;min-width:24px;}',
    '.mc-choice-text{flex:1;line-height:1.6;}',
    '.mc-action-row{display:flex;gap:10px;margin-top:12px;}',
    '.answer-input{flex:1;min-width:220px;border:1px solid #c8d1dc;border-radius:10px;padding:12px 14px;font-size:14px;outline:none;background:#fff;}',
    '.feedback{margin-top:10px;min-height:22px;font-size:13px;font-weight:700;}',
    '.answer-reveal{margin-top:10px;padding:10px 12px;border-radius:10px;background:#fff8e6;border:1px solid #f2d27c;font-size:13px;line-height:1.5;}',
    '.teacher-answer-box{margin-top:10px;padding:10px 12px;border-radius:10px;background:#f9fbfe;border:1px dashed #b8c7db;font-size:13px;line-height:1.6;}',
    '.hint-line{margin-top:6px;color:#5b6b82;}',
    '.hidden{display:none;}',
    '.objective-text,.summary-text,.concept-summary,.sub-card,.example-box,.misconception-box,.practice-question,.vocab-definition,.teacher-answer-box,.answer-reveal,.hint-line{white-space:pre-line;overflow-wrap:anywhere;word-break:break-word;}',
    '.option-preview{margin:10px 0 12px;padding:10px 12px;background:var(--template-subtle-bg);border:1px solid var(--template-subtle-border);border-radius:12px;}',
    '.option-line{margin:4px 0;font-size:13px;line-height:1.5;}',
    '.print-answer-space{height:28px;border-bottom:1px solid #7a869a;margin-top:18px;}',
    '.print-answer-space.lines-2{height:56px;}',
    '.editable.active-edit{outline:2px dashed var(--template-accent);outline-offset:2px;background:var(--template-soft-bg);cursor:text;}',
    '.lesson-sheet .btn-primary{background:var(--template-accent);color:#fff;}',
    '.lesson-sheet .btn-light{background:var(--template-soft-bg);color:var(--template-accent-dark);border:1px solid var(--template-soft-border);}',
    '@media (max-width: 760px){.lesson-meta,.vocab-grid,.review-grid,.quiz-meta-lines,.guided-stage-grid,.review-list,.remediation-grid,.enrichment-grid{grid-template-columns:1fr;}.lesson-sheet[data-template="REVIEW"] .compact-vocab-list{columns:1;}}'
  ].join('');
}


function buildLessonGuide(formData) {
  const request = sanitizeRequest_(formData);
  const lesson = generateLessonJson_(request);

  return {
    ok: true,
    request: request,
    lesson: lesson,
    html: renderLessonHtml_(lesson, {
      forPrint: false,
      copyMode: request.previewMode,
      layoutMode: request.layoutMode,
      headerInfo: getHeaderInfoFromRequest_(request)
    })
  };
}

function renderLessonPreview(payload) {
  if (!payload || !payload.lesson) {
    throw new Error('Missing lesson data for preview rendering.');
  }

  const lesson = normalizeLesson_(payload.lesson, payload.request || {});
  validateLesson_(lesson, payload.request || {});

  return {
    html: renderLessonHtml_(lesson, {
      forPrint: false,
      copyMode: payload.copyMode || 'STUDENT',
      layoutMode: (payload.request && payload.request.layoutMode) || DEFAULT_LAYOUT_MODE,
      headerInfo: getHeaderInfoFromRequest_(payload.request || {})
    })
  };
}

function getPrintableLessonHtml(payload) {
  if (!payload || !payload.lesson) {
    throw new Error('Missing lesson data for print export.');
  }

  const request = sanitizeRequest_(payload.request || {});
  const lesson = normalizeLesson_(payload.lesson, request);
  validateLesson_(lesson, request);

  const copyMode = String(payload.copyMode || request.previewMode || 'STUDENT').toUpperCase();
  const layoutMode = payload.layoutMode || request.layoutMode || DEFAULT_LAYOUT_MODE;
  const headerInfo = getHeaderInfoFromRequest_(request);

  const bodyHtml = renderLessonHtml_(lesson, {
    forPrint: true,
    copyMode: copyMode,
    layoutMode: layoutMode,
    headerInfo: headerInfo
  });

  const suggestedFilename =
    String(request.topic || lesson.topic || 'Lesson Guide').trim() +
    ' - ' +
    prettyEnum_(request.templateType || lesson.templateType || 'CONCEPT');

  return {
    filename: suggestedFilename + '.pdf',
    html: buildPrintableDocument_(bodyHtml, {
      layoutMode: layoutMode,
      documentTitle: suggestedFilename,
      footerText: '© ' + new Date().getFullYear() + ' Generated using ' + APP_TITLE + '.'
    })
  };
}

/***************************************
 * 3. GEMINI CONFIG
 ***************************************/
function getGeminiConfig_() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('GEMINI_API_KEY');
  const model = props.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;

  if (!apiKey) {
    throw new Error('Missing GEMINI_API_KEY in Script Properties.');
  }

  return { apiKey: apiKey, model: model };
}

/***************************************
 * 4. GEMINI GENERATION
 ***************************************/
function generateLessonJson_(request) {
  const config = getGeminiConfig_();
  const apiKey = config.apiKey;
  const model = config.model;

  const schema = getLessonSchema_(request);
  const prompt = buildPrompt_(request);
  const url =
    'https://generativelanguage.googleapis.com/v1beta/' +
    model +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);

  let lastError = 'Unknown generation error.';
  let outputText = '';

  for (let attempt = 1; attempt <= 2; attempt++) {
    const payload = {
      contents: [
        {
          parts: [
            {
              text:
                prompt +
                (attempt === 2
                  ? '\n\nReturn JSON only. No markdown fences. No commentary. Every string must be properly quoted and escaped. Follow the schema exactly.'
                  : '')
            }
          ]
        }
      ],
      generationConfig: {
        maxOutputTokens: 8192,
        responseMimeType: 'application/json',
        responseJsonSchema: schema
      }
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const status = response.getResponseCode();
    const text = response.getContentText();

    if (status !== 200) {
      lastError = extractApiError_(text) || ('Gemini API request failed with status ' + status + '.');
      continue;
    }

    try {
      const raw = JSON.parse(text);
      outputText = extractCandidateText_(raw);

      if (!outputText) {
        lastError = 'Gemini returned an empty response.';
        continue;
      }

      const cleanedJson = extractJsonText_(outputText);
      const parsed = JSON.parse(cleanedJson);
      const lesson = normalizeLesson_(parsed, request);
      validateLesson_(lesson, request);
      return lesson;
    } catch (err) {
      try {
        const repaired = repairLessonJson_(outputText || text, schema, apiKey, model);
        const lesson = normalizeLesson_(repaired, request);
        validateLesson_(lesson, request);
        return lesson;
      } catch (repairErr) {
        lastError =
          'Could not parse or validate lesson JSON: ' +
          err.message +
          ' | Repair attempt failed: ' +
          repairErr.message;
      }
    }
  }

  throw new Error(lastError);
}

function buildPrompt_(request) {
  const subjectGuidance = getSubjectGuidance_(request.subject);
  const templateGuidance = getTemplateGuidance_(request.templateType);
  const templateFieldBehavior = getTemplateFieldBehavior_(request.templateType);
  const localContextGuidance = getLocalContextGuidance_();
  const difficultyGuidance = getDifficultyGuidance_(request.difficulty);
  const practiceDesignGuidance = getPracticeDesignGuidance_(request.templateType, request.difficulty);

  return [
    'You are an expert instructional content generator for classroom teachers.',
    'Create a classroom-ready lesson guide in JSON only.',
    '',
    'Required output:',
    '- Return valid JSON only.',
    '- Fill all schema fields completely and meaningfully.',
    '- Ensure every field is classroom-ready, specific, and useful.',
    '',
    'Key concept rules:',
    '- Each key concept must include: heading, summary, formula, symbolMeaning, workedExample, misconception.',
    '- For non-math topics, formula may be an empty string only if no formula is appropriate.',
    '- symbolMeaning must explain symbols or important vocabulary used in the formula or concept.',
    '- workedExample must begin with a short problem or context before the steps.',
    '- For Mathematics, present workedExample in this order: Problem, Step 1, Step 2, Step 3, and continue with additional steps if needed, then Answer.',
    '- Provide a logical step-by-step sequence with at least 3 steps when the topic requires procedural solving.',
    '- For simpler topics, keep the solution short but clearly sequenced.',
    '- Keep the Problem and Step explanations in plain readable text.',
    '- Use KaTeX only for the actual calculation line, not for every number inside the sentence.',
    '- Place each step on its own new line using the format "Step 1:", "Step 2:", and so on.',
    '- End workedExample with a final line in the format "Answer: ...".',
    '- misconception must be practical and classroom-relevant.',
    '',
    'Practice item rules:',
    '- Create EXACTLY ' + request.itemCount + ' practice items.',
    '- Each item must include: type, question, options, acceptedAnswers, hint.',
    '- Allowed types: short-answer, multiple-choice, true-false.',
    '- For short-answer, options must be an empty array.',
    '- For multiple-choice, provide exactly 4 options.',
    '- For true-false, options must be ["True","False"].',
    '- acceptedAnswers must contain at least one expected answer string.',
    '- The practice items must match the selected template type in tone, sequence, and difficulty progression.',
    '- Keep all questions age-appropriate, clear, and answerable.',
    '',
    'Formatting rules:',
    '- In Mathematics lessons, use display KaTeX \\[...\\] in the formula field for every formula or computation line.',
    '- In Mathematics lessons, use inline KaTeX \\(...\\) for symbolic math inside sentences, such as \\(A_{\\text{total}}\\), \\(A_1\\), \\(A_2\\), \\(x^2\\), and \\(\\frac{3}{4}\\).',
    '- Any expression with variables, subscripts, superscripts, fractions, exponents, roots, operators, or equal signs must be written in KaTeX.',
    '- Never write plain-text math notation such as A_total, A1, x2, y3, or 1/2 when mathematical notation is intended.',
    '- Keep ordinary explanatory words in plain text, but wrap every symbolic part in KaTeX.',
    '- Put each display calculation on its own line.',
    '- Use \\[...\\] for standalone calculations or important formulas.',
    '- Do not use markdown fences.',
    '- Return JSON only.',
    '',
    'Instructional settings:',
    'Subject: ' + request.subject,
    'Grade Level: ' + request.gradeLevel,
    'Topic: ' + request.topic,
    'Template Type: ' + request.templateType,
    'Difficulty: ' + request.difficulty,
    'Objective/Competency: ' + (request.objective || 'Use the topic to infer a strong lesson objective.'),
    'Use real-world context: ' + (request.includeRealWorld ? 'Yes' : 'No'),
    'Include success criteria: ' + (request.includeSuccessCriteria ? 'Yes' : 'No'),
    'Include vocabulary: ' + (request.includeVocabulary ? 'Yes' : 'No'),
    'Include misconceptions: ' + (request.includeMisconceptions ? 'Yes' : 'No'),
    '',
    'Local classroom context:',
    localContextGuidance,
    '',
    'Subject-specific guidance:',
    subjectGuidance,
    '',
    'Template-specific guidance:',
    templateGuidance,
    '',
    'Template-specific field behavior:',
    templateFieldBehavior,
    '',
    'Difficulty-specific guidance:',
    difficultyGuidance,
    '',
    'Practice design guidance:',
practiceDesignGuidance,
'',
    'Quality expectations:',
    '- lessonSummary must be genuinely useful for classroom instruction and must fit the selected template.',
    '- successCriteria must be student-friendly and measurable.',
    '- vocabulary terms must be age-appropriate.',
    '- teacherNotes must include teaching tips, pacing advice, or likely misconceptions.',
    '- Ensure the lesson is useful, specific, and classroom-ready rather than generic.',
    '- Make the tone, section content, and practice style clearly match the selected template.'
  ].join('\n');
}

function getSubjectGuidance_(subject) {
  const s = String(subject || '').toLowerCase();

  if (s.indexOf('math') !== -1) {
    return [
      '- This is a Mathematics lesson.',
      '- Provide 4 to 6 key concepts.',
      '- Every key concept must include a non-empty formula in KaTeX if the topic uses a rule, property, equation, or computational relationship.',
      '- Define all symbols clearly in symbolMeaning.',
      '- workedExample must start with a real problem or context before showing steps.',
      '- Use a short word problem, numerical situation, or visual interpretation prompt before the solution steps.',
      '- workedExample must show step-by-step reasoning on separate lines.',
      '- In workedExample and symbolMeaning, every symbolic reference must use KaTeX, for example \\(A_{\\text{total}}\\), \\(A_1\\), and \\(A_2\\).',
      '- End each workedExample with a final answer line.',
      '- misconception must address a likely student error.',
      '- Practice items should progress from recall to application when possible.'
    ].join('\n');
  }

  if (s.indexOf('science') !== -1) {
    return [
      '- Emphasize scientific accuracy, observation, explanation, and cause-and-effect.',
      '- Use concrete real-world examples and clear terminology.'
    ].join('\n');
  }

  if (s.indexOf('english') !== -1) {
    return [
      '- Emphasize grammar, vocabulary, comprehension, and clear expression.',
      '- Use age-appropriate examples and concise explanations.'
    ].join('\n');
  }

  if (s.indexOf('filipino') !== -1) {
    return [
      '- Gumamit ng malinaw, wasto, at angkop na wikang Filipino.',
      '- Tiyakin na ang mga paliwanag at halimbawa ay angkop sa antas ng mag-aaral.'
    ].join('\n');
  }

  if (s.indexOf('araling') !== -1 || s.indexOf('panlipunan') !== -1 || s === 'ap') {
    return [
      '- Emphasize chronology, geography, civics, source-based understanding, or social interpretation as needed.',
      '- Use concise factual examples and student-friendly wording.'
    ].join('\n');
  }

  return [
    '- Keep the lesson accurate, clear, and age-appropriate.',
    '- Balance explanation, examples, and guided practice.'
  ].join('\n');
}

function getTemplateGuidance_(templateType) {
  const t = String(templateType || DEFAULT_TEMPLATE).toUpperCase();

  if (t === 'GUIDED_PRACTICE') {
    return [
      '- Purpose: create a scaffolded lesson that moves from explanation to supported example to independent practice.',
      '- Tone: clear, supportive, teacher-guided, and step-by-step.',
      '- lessonSummary should describe the learning path and what students will do during guided practice.',
      '- keyConcepts should be concise, well-sequenced, and easy to act on.',
      '- Keep sections brief and actionable, focusing on student execution.',
      '- workedExample must be especially strong, explicit, and easy to follow.',
      '- Practice items must progress from easier to harder.',
      '- Early practice items should be direct and highly accessible.',
      '- Use language that supports confidence, clarity, and gradual release.'
    ].join('\n');
  }

  if (t === 'REVIEW') {
    return [
      '- Purpose: create a compact review sheet for recalling and checking previously learned ideas.',
      '- Tone: concise, efficient, and review-oriented.',
      '- lessonSummary should be brief and focused on the main ideas students need to remember.',
      '- keyConcept summaries should be short and written as quick reminders.',
      '- symbolMeaning should function as a memory aid or review cue.',
      '- workedExample should be brief and selective rather than fully instructional.',
      '- Practice items should provide representative mixed review across the topic.',
      '- Include both recall and straightforward application.',
      '- Keep explanations compact so the sheet feels fast to scan and use.',
      '- Do not turn the review into a full direct-instruction lesson.'
    ].join('\n');
  }

  if (t === 'QUIZ') {
    return [
      '- Purpose: create an assessment-style lesson guide that supports a quiz format.',
      '- Tone: formal, neutral, direct, and assessment-oriented.',
      '- lessonSummary should be minimal and concise because the main emphasis is assessment.',
      '- keyConcepts may still be included for schema completeness, but keep them compact and reference-style.',
      '- Practice items must read like actual quiz items with clear, answerable wording.',
      '- Use direct, objective testing language such as "Solve", "Identify", "Choose", "Compare", or "Determine".',
      '- Questions should feel independent and test-ready.',
      '- Hints should stay short and practical.',
      '- Use precise wording and avoid ambiguity.',
      '- Do not include tutorial phrasing inside the quiz questions.'
    ].join('\n');
  }

  if (t === 'REMEDIATION') {
    return [
      '- Purpose: support learners who need simpler explanations, smaller steps, and more guided understanding.',
      '- Tone: supportive, clear, patient, and highly accessible.',
      '- Use short sentences, explicit wording, and concrete examples.',
      '- lessonSummary should clearly explain the focus in a reassuring and easy-to-follow way.',
      '- keyConcept summaries should be broken into manageable ideas.',
      '- workedExample must be highly explicit and carefully sequenced.',
      '- misconception should identify a very common student error in practical classroom language.',
      '- Practice items must begin with very accessible tasks before moving to slightly more independent work.',
      '- Build difficulty gradually and keep cognitive load low in the earlier items.'
    ].join('\n');
  }

  if (t === 'ENRICHMENT') {
    return [
      '- Purpose: extend learning beyond standard mastery through deeper thinking, reasoning, and transfer.',
      '- Tone: intellectually engaging, challenging, and thought-provoking.',
      '- lessonSummary should frame the topic as an opportunity to think more deeply or apply ideas in richer ways.',
      '- keyConcept summaries should emphasize patterns, relationships, reasoning, or transfer when appropriate.',
      '- workedExample should model deeper thinking, not just routine procedure.',
      '- Practice items should become more demanding as they progress.',
      '- Ensure at least 50% of practice items require multi-step reasoning, pattern recognition, comparison, or real-world application.',
      '- Later items should involve transfer to new or less familiar situations.',
      '- Increase the thinking demand, not just the size of the numbers.'
    ].join('\n');
  }

  return [
    '- Purpose: create a full concept lesson for classroom instruction.',
    '- Tone: clear, explanatory, and classroom-ready.',
    '- lessonSummary should provide a strong teaching overview.',
    '- keyConcepts should support direct instruction with clear explanation and useful examples.',
    '- workedExample should be complete, readable, and instructionally helpful.',
    '- Practice items should move from understanding toward application.',
    '- Balance explanation, structure, and classroom usefulness.'
  ].join('\n');
}

function getTemplateGuidance_(templateType) {
  const t = String(templateType || DEFAULT_TEMPLATE).toUpperCase();

  if (t === 'GUIDED_PRACTICE') {
    return [
      '- Purpose: create a scaffolded lesson that moves from explanation to supported example to independent practice.',
      '- Tone: clear, supportive, teacher-guided, and step-by-step.',
      '- lessonSummary should describe the learning path and what students will do during guided practice.',
      '- keyConcepts should be concise, well-sequenced, and easy to act on.',
      '- Keep sections brief and actionable, focusing on student execution.',
      '- workedExample must be especially strong, explicit, and easy to follow.',
      '- Practice items must progress from easier to harder.',
      '- Early practice items should be direct and highly accessible.',
      '- Use language that supports confidence, clarity, and gradual release.'
    ].join('\n');
  }

  if (t === 'REVIEW') {
    return [
      '- Purpose: create a compact review sheet for recalling and checking previously learned ideas.',
      '- Tone: concise, efficient, and review-oriented.',
      '- lessonSummary should be brief and focused on the main ideas students need to remember.',
      '- keyConcept summaries should be short and written as quick reminders.',
      '- symbolMeaning should function as a memory aid or review cue.',
      '- workedExample should be brief and selective rather than fully instructional.',
      '- Practice items should provide representative mixed review across the topic.',
      '- Include both recall and straightforward application.',
      '- Keep explanations compact so the sheet feels fast to scan and use.',
      '- Do not turn the review into a full direct-instruction lesson.'
    ].join('\n');
  }

  if (t === 'QUIZ') {
    return [
      '- Purpose: create an assessment-style lesson guide that supports a quiz format.',
      '- Tone: formal, neutral, direct, and assessment-oriented.',
      '- lessonSummary should be minimal and concise because the main emphasis is assessment.',
      '- keyConcepts may still be included for schema completeness, but keep them compact and reference-style.',
      '- Practice items must read like actual quiz items with clear, answerable wording.',
      '- Use direct, objective testing language such as "Solve", "Identify", "Choose", "Compare", or "Determine".',
      '- Questions should feel independent and test-ready.',
      '- Hints should stay short and practical.',
      '- Use precise wording and avoid ambiguity.',
      '- Do not include tutorial phrasing inside the quiz questions.'
    ].join('\n');
  }

  if (t === 'REMEDIATION') {
    return [
      '- Purpose: support learners who need simpler explanations, smaller steps, and more guided understanding.',
      '- Tone: supportive, clear, patient, and highly accessible.',
      '- Use short sentences, explicit wording, and concrete examples.',
      '- lessonSummary should clearly explain the focus in a reassuring and easy-to-follow way.',
      '- keyConcept summaries should be broken into manageable ideas.',
      '- workedExample must be highly explicit and carefully sequenced.',
      '- misconception should identify a very common student error in practical classroom language.',
      '- Practice items must begin with very accessible tasks before moving to slightly more independent work.',
      '- Build difficulty gradually and keep cognitive load low in the earlier items.'
    ].join('\n');
  }

  if (t === 'ENRICHMENT') {
    return [
      '- Purpose: extend learning beyond standard mastery through deeper thinking, reasoning, and transfer.',
      '- Tone: intellectually engaging, challenging, and thought-provoking.',
      '- lessonSummary should frame the topic as an opportunity to think more deeply or apply ideas in richer ways.',
      '- keyConcept summaries should emphasize patterns, relationships, reasoning, or transfer when appropriate.',
      '- workedExample should model deeper thinking, not just routine procedure.',
      '- Practice items should become more demanding as they progress.',
      '- Ensure at least 50% of practice items require multi-step reasoning, pattern recognition, comparison, or real-world application.',
      '- Later items should involve transfer to new or less familiar situations.',
      '- Increase the thinking demand, not just the size of the numbers.'
    ].join('\n');
  }

  return [
    '- Purpose: create a full concept lesson for classroom instruction.',
    '- Tone: clear, explanatory, and classroom-ready.',
    '- lessonSummary should provide a strong teaching overview.',
    '- keyConcepts should support direct instruction with clear explanation and useful examples.',
    '- workedExample should be complete, readable, and instructionally helpful.',
    '- Practice items should move from understanding toward application.',
    '- Balance explanation, structure, and classroom usefulness.'
  ].join('\n');
}

function getTemplateFieldBehavior_(templateType) {
  const t = String(templateType || DEFAULT_TEMPLATE).toUpperCase();

  if (t === 'GUIDED_PRACTICE') {
    return [
      '- lessonSummary must describe the learning flow and what students will do during the lesson.',
      '- keyConcept summaries must be brief, clear, and sequenced for guided instruction.',
      '- workedExample must be especially strong and highly usable for the PRACTICE stage.',
      '- Practice items must start with accessible items and increase gradually in difficulty.',
      '- Use a helpful mix of short-answer, multiple-choice, and true-false when appropriate.',
      '- The first third of the practice items should feel very approachable.'
    ].join('\n');
  }

  if (t === 'REVIEW') {
    return [
      '- lessonSummary must be brief and focused on the most important ideas to remember.',
      '- keyConcept summaries must read like quick reminders, not long explanations.',
      '- symbolMeaning should act like a review cue or memory aid.',
      '- workedExample should be shorter and more selective than in a concept lesson.',
      '- Practice items should give representative mixed review across the topic.',
      '- Keep item wording concise and easy to scan.'
    ].join('\n');
  }

  if (t === 'QUIZ') {
    return [
      '- lessonSummary must be very short and assessment-oriented.',
      '- keyConcepts should remain compact because they mainly support teacher use and schema completeness.',
      '- Practice items must function like actual assessment items.',
      '- Questions must be directly answerable and independently readable.',
      '- Use direct command wording such as "Solve", "Find", "Choose", "Identify", "Compare", or "Determine".',
      '- Hints must stay short and practical.'
    ].join('\n');
  }

  if (t === 'REMEDIATION') {
    return [
      '- lessonSummary must be reassuring, clear, and easy to follow.',
      '- keyConcept summaries must be broken into smaller and simpler ideas.',
      '- workedExample must be highly explicit and carefully sequenced.',
      '- misconception must name a very common student mistake in plain classroom language.',
      '- Practice items must begin with very accessible items before moving to slightly more independent work.',
      '- The first half of the practice items should emphasize clarity and confidence-building.'
    ].join('\n');
  }

  if (t === 'ENRICHMENT') {
    return [
      '- lessonSummary must frame the topic as deeper thinking, richer application, or transfer.',
      '- keyConcept summaries should emphasize patterns, relationships, and reasoning.',
      '- workedExample should model deeper thought, not only routine computation.',
      '- Practice items must become more demanding as they progress.',
      '- At least half of the practice items should require multi-step reasoning, comparison, pattern recognition, or real-world transfer.',
      '- The later items should feel more challenging than standard on-level practice.'
    ].join('\n');
  }

  return [
    '- lessonSummary must provide a strong teaching overview.',
    '- keyConcept summaries should support direct instruction with clear explanation.',
    '- workedExample should be complete, readable, and instructionally helpful.',
    '- Practice items should move from understanding toward application.',
    '- Maintain a balanced mix of explanation and practice.'
  ].join('\n');
}

function getLocalContextGuidance_() {
  return [
    '- Use Philippine classroom context when examples involve names, money, measurement, or everyday situations.',
    '- Prefer pesos (₱), metric units, and familiar school or community-based situations when appropriate.',
    '- Use age-appropriate local names and realistic classroom-friendly contexts.'
  ].join('\n');
}

function getDifficultyGuidance_(difficulty) {
  const d = String(difficulty || 'ON_LEVEL').toUpperCase();

  if (d === 'BELOW_LEVEL') {
    return [
      '- Use simpler wording, shorter sentences, and more concrete examples.',
      '- Keep explanations highly accessible and reduce unnecessary abstraction.',
      '- Break ideas into smaller and more manageable steps.',
      '- Practice items should begin with direct, confidence-building questions.',
      '- Use easier early items and smaller jumps in difficulty.',
      '- Prefer clarity, support, and strong scaffolding over complexity.'
    ].join('\n');
  }

  if (d === 'ABOVE_LEVEL') {
    return [
      '- Use richer reasoning, stronger conceptual connections, and more demanding application.',
      '- Allow greater complexity while keeping the lesson age-appropriate.',
      '- Practice items should include more multi-step reasoning and less routine repetition.',
      '- Later items should require transfer, comparison, justification, or deeper analysis when appropriate.',
      '- Increase the thinking demand, not just the size of the numbers or the length of the wording.',
      '- Keep the content challenging but still clear and classroom-usable.'
    ].join('\n');
  }

  return [
    '- Use grade-appropriate wording, pacing, and conceptual demand.',
    '- Balance clarity, independence, and reasonable challenge.',
    '- Practice items should move through a standard progression from understanding to application.',
    '- Maintain normal grade-level expectations for vocabulary, explanation, and reasoning.'
  ].join('\n');
}

function getPracticeDesignGuidance_(templateType, difficulty) {
  const t = String(templateType || DEFAULT_TEMPLATE).toUpperCase();
  const d = String(difficulty || 'ON_LEVEL').toUpperCase();
  const rules = [];

  if (t === 'GUIDED_PRACTICE') {
    rules.push(
      '- For GUIDED_PRACTICE, sequence practice items from most accessible to more independent.',
      '- The early items should feel supported and direct.',
      '- Later items may require more independence, but the progression must stay smooth.'
    );
  } else if (t === 'REVIEW') {
    rules.push(
      '- For REVIEW, use representative mixed review across the topic.',
      '- Include a balanced spread of recall and straightforward application.',
      '- Keep item wording concise and easy to scan.'
    );
  } else if (t === 'QUIZ') {
    rules.push(
      '- For QUIZ, all practice items must function as assessment items rather than guided activities.',
      '- Use neutral, direct, test-style wording.',
      '- Keep the sequence balanced and formal.'
    );
  } else if (t === 'REMEDIATION') {
    rules.push(
      '- For REMEDIATION, begin with highly accessible items and increase independence gradually.',
      '- The early items should reduce cognitive load and build confidence.',
      '- Keep wording very clear and concrete.'
    );
  } else if (t === 'ENRICHMENT') {
    rules.push(
      '- For ENRICHMENT, make later items more demanding than earlier ones.',
      '- Include non-routine application, pattern recognition, transfer, comparison, or deeper reasoning.',
      '- Ensure the practice is challenging through thinking demand, not just bigger numbers.'
    );
  } else {
    rules.push(
      '- For CONCEPT, move from understanding toward application in a balanced way.',
      '- Start with direct understanding items, then progress toward application.'
    );
  }

  if (d === 'BELOW_LEVEL') {
    rules.push(
      '- Because the difficulty is BELOW_LEVEL, use shorter, clearer, and more direct questions.',
      '- The first half of the items should be especially accessible and confidence-building.',
      '- Keep reasoning demand moderate and avoid large jumps in complexity.'
    );
  } else if (d === 'ABOVE_LEVEL') {
    rules.push(
      '- Because the difficulty is ABOVE_LEVEL, increase reasoning demand across the set.',
      '- Include more multi-step, comparison, transfer, or justification items, especially in the later portion.',
      '- Raise the conceptual challenge while keeping wording age-appropriate and clear.'
    );
  } else {
    rules.push(
      '- Because the difficulty is ON_LEVEL, maintain a standard grade-level progression from understanding to application.',
      '- Keep the overall challenge balanced and appropriate for the selected grade level.'
    );
  }

  return rules.join('\n');
}

/***************************************
 * 5. SCHEMA
 ***************************************/
function getLessonSchema_(request) {
  return {
    type: 'object',
    additionalProperties: false,
    required: [
      'title',
      'subject',
      'gradeLevel',
      'topic',
      'templateType',
      'difficulty',
      'objective',
      'successCriteria',
      'vocabulary',
      'lessonSummary',
      'keyConcepts',
      'practiceItems',
      'teacherNotes'
    ],
    properties: {
      title: { type: 'string', minLength: 3 },
      subject: { type: 'string', minLength: 2 },
      gradeLevel: { type: 'string', minLength: 1 },
      topic: { type: 'string', minLength: 2 },
      templateType: { type: 'string', minLength: 3 },
      difficulty: { type: 'string', minLength: 2 },
      objective: { type: 'string', minLength: 10 },
      successCriteria: {
        type: 'array',
        minItems: request.includeSuccessCriteria ? 3 : 0,
        maxItems: 6,
        items: { type: 'string', minLength: 5 }
      },
      vocabulary: {
        type: 'array',
        minItems: request.includeVocabulary ? 3 : 0,
        maxItems: 8,
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['term', 'definition'],
          properties: {
            term: { type: 'string', minLength: 2 },
            definition: { type: 'string', minLength: 5 }
          }
        }
      },
      lessonSummary: { type: 'string', minLength: 60 },
      keyConcepts: {
        type: 'array',
        minItems: isMathSubject_(request.subject) ? 4 : 3,
        maxItems: 6,
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['heading', 'summary', 'formula', 'symbolMeaning', 'workedExample', 'misconception'],
          properties: {
            heading: { type: 'string', minLength: 2 },
            summary: { type: 'string', minLength: 25 },
            formula: { type: 'string' },
            symbolMeaning: { type: 'string', minLength: 5 },
            workedExample: { type: 'string', minLength: 15 },
            misconception: { type: 'string', minLength: request.includeMisconceptions ? 8 : 0 }
          }
        }
      },
      practiceItems: {
        type: 'array',
        minItems: request.itemCount,
        maxItems: request.itemCount,
        items: {
          type: 'object',
          additionalProperties: false,
          required: ['type', 'question', 'options', 'acceptedAnswers', 'hint'],
          properties: {
            type: { type: 'string', minLength: 1 },
            question: { type: 'string', minLength: 5 },
            options: {
              type: 'array',
              minItems: 0,
              maxItems: 4,
              items: { type: 'string', minLength: 1 }
            },
            acceptedAnswers: {
              type: 'array',
              minItems: 1,
              maxItems: 6,
              items: { type: 'string', minLength: 1 }
            },
            hint: { type: 'string', minLength: 1 }
          }
        }
      },
      teacherNotes: {
        type: 'array',
        minItems: 3,
        maxItems: 8,
        items: { type: 'string', minLength: 8 }
      }
    }
  };
}

/***************************************
 * 6. REQUEST / LESSON CLEANING
 ***************************************/
function sanitizeRequest_(formData) {
  const data = formData || {};
  const cleaned = {
    subject: String(data.subject || '').trim(),
    gradeLevel: String(data.gradeLevel || '').trim(),
    topic: String(data.topic || '').trim(),
    objective: normalizeDisplayText_(data.objective || ''),
    previewMode: String(data.previewMode || 'STUDENT').trim().toUpperCase(),
    layoutMode: String(data.layoutMode || DEFAULT_LAYOUT_MODE).trim().toUpperCase(),
    templateType: String(data.templateType || DEFAULT_TEMPLATE).trim().toUpperCase(),
    difficulty: String(data.difficulty || 'ON_LEVEL').trim().toUpperCase(),
    itemCount: toBoundedInteger_(data.itemCount, DEFAULT_ITEM_COUNT, 8, 20),
    includeSuccessCriteria: toBoolean_(data.includeSuccessCriteria, true),
    includeVocabulary: toBoolean_(data.includeVocabulary, true),
    includeMisconceptions: toBoolean_(data.includeMisconceptions, true),
    includeRealWorld: toBoolean_(data.includeRealWorld, true),
    schoolName: String(data.schoolName || '').trim(),
    teacherName: String(data.teacherName || '').trim(),
    quarter: String(data.quarter || '').trim(),
    customHeader: String(data.customHeader || '').trim()
  };

  if (!cleaned.subject) throw new Error('Subject is required.');
  if (!cleaned.gradeLevel) throw new Error('Grade level is required.');
  if (!cleaned.topic) throw new Error('Topic is required.');

  if (['STUDENT', 'TEACHER'].indexOf(cleaned.previewMode) === -1) cleaned.previewMode = 'STUDENT';
  if (['COMPACT', 'STANDARD', 'SPACIOUS'].indexOf(cleaned.layoutMode) === -1) cleaned.layoutMode = DEFAULT_LAYOUT_MODE;
  if (['CONCEPT', 'GUIDED_PRACTICE', 'REVIEW', 'QUIZ', 'REMEDIATION', 'ENRICHMENT'].indexOf(cleaned.templateType) === -1) cleaned.templateType = DEFAULT_TEMPLATE;
  if (['BELOW_LEVEL', 'ON_LEVEL', 'ABOVE_LEVEL'].indexOf(cleaned.difficulty) === -1) cleaned.difficulty = 'ON_LEVEL';

  return cleaned;
}

function normalizeLesson_(lesson, fallbackRequest) {
  const base = lesson || {};
  const req = fallbackRequest || {};
  const expectedItemCount = toBoundedInteger_(req.itemCount, DEFAULT_ITEM_COUNT, 8, 20);

  const keyConcepts = Array.isArray(base.keyConcepts) ? base.keyConcepts : [];
  const practiceItems = Array.isArray(base.practiceItems) ? base.practiceItems : [];
  const teacherNotes = Array.isArray(base.teacherNotes) ? base.teacherNotes : [];
  const successCriteria = Array.isArray(base.successCriteria) ? base.successCriteria : [];
  const vocabulary = Array.isArray(base.vocabulary) ? base.vocabulary : [];

  return {
    title: String(base.title || req.topic || 'Lesson Guide').trim(),
    subject: String(base.subject || req.subject || '').trim(),
    gradeLevel: String(base.gradeLevel || req.gradeLevel || '').trim(),
    topic: String(base.topic || req.topic || '').trim(),
    templateType: String(base.templateType || req.templateType || DEFAULT_TEMPLATE).trim(),
    difficulty: String(base.difficulty || req.difficulty || 'ON_LEVEL').trim(),
    objective: String(base.objective || req.objective || '').trim(),
    successCriteria: successCriteria.map(function (item) {
      return String(item || '').trim();
    }).filter(Boolean),
    vocabulary: vocabulary.map(function (item) {
      return {
        term: String((item && item.term) || '').trim(),
        definition: String((item && item.definition) || '').trim()
      };
    }).filter(function (item) {
      return item.term && item.definition;
    }),
    lessonSummary: normalizeDisplayText_(base.lessonSummary || ''),
    keyConcepts: keyConcepts.map(function (item) {
      return {
        heading: String((item && item.heading) || '').trim(),
        summary: normalizeDisplayText_((item && item.summary) || ''),
        formula: String((item && item.formula) || '').trim(),
        symbolMeaning: normalizeDisplayText_((item && item.symbolMeaning) || ''),
        workedExample: normalizeDisplayText_((item && item.workedExample) || ''),
        misconception: normalizeDisplayText_((item && item.misconception) || '')
      };
    }).filter(function (item) {
      return item.heading && item.summary && item.workedExample;
    }),
    practiceItems: practiceItems.map(function (item) {
      const type = normalizeItemType_(item && item.type);
      let options = Array.isArray(item && item.options) ? item.options : [];
      const acceptedAnswers = Array.isArray(item && item.acceptedAnswers)
        ? item.acceptedAnswers
        : ((item && item.answer) ? [String(item.answer)] : []);

      if (type === 'true-false' && options.length === 0) {
        options = ['True', 'False'];
      }

      return {
        type: type,
        question: String((item && item.question) || '').trim(),
        options: options.map(function (o) { return String(o || '').trim(); }).filter(Boolean),
        acceptedAnswers: acceptedAnswers.map(function (a) { return String(a || '').trim(); }).filter(Boolean),
        hint: String((item && item.hint) || '').trim()
      };
    }).filter(function (item) {
      return item.question && item.acceptedAnswers.length;
    }).slice(0, expectedItemCount),
    teacherNotes: teacherNotes.map(function (note) {
      return String(note || '').trim();
    }).filter(Boolean)
  };
}

function normalizeItemType_(value) {
  const t = String(value || '').toLowerCase().trim();
  if (t === 'multiple-choice' || t === 'multiple choice' || t === 'mcq') return 'multiple-choice';
  if (t === 'true-false' || t === 'true/false' || t === 'tf') return 'true-false';
  return 'short-answer';
}

function validateLesson_(lesson, request) {
  const req = request || {};
  const expectedItemCount = toBoundedInteger_(req.itemCount, DEFAULT_ITEM_COUNT, 8, 20);

  if (!lesson.title) throw new Error('Lesson title is missing.');
  if (!lesson.subject) throw new Error('Lesson subject is missing.');
  if (!lesson.gradeLevel) throw new Error('Lesson grade level is missing.');
  if (!lesson.topic) throw new Error('Lesson topic is missing.');
  if (!lesson.lessonSummary) throw new Error('Lesson summary is missing.');
  if (!lesson.objective) throw new Error('Lesson objective is missing.');

  if (!Array.isArray(lesson.keyConcepts) || lesson.keyConcepts.length < (isMathSubject_(lesson.subject) ? 4 : 3)) {
    throw new Error('Lesson must contain enough key concepts for the selected subject.');
  }

  if (!Array.isArray(lesson.practiceItems) || lesson.practiceItems.length !== expectedItemCount) {
    throw new Error('Lesson must contain exactly ' + expectedItemCount + ' practice items.');
  }

  if (req.includeSuccessCriteria && (!Array.isArray(lesson.successCriteria) || lesson.successCriteria.length < 3)) {
    throw new Error('Lesson must contain at least 3 success criteria.');
  }

  if (req.includeVocabulary && (!Array.isArray(lesson.vocabulary) || lesson.vocabulary.length < 3)) {
    throw new Error('Lesson must contain at least 3 vocabulary entries.');
  }

  lesson.keyConcepts.forEach(function (item, index) {
    if (!item.heading || !item.summary || !item.workedExample) {
      throw new Error('Key concept ' + (index + 1) + ' is incomplete.');
    }

    if (isMathSubject_(lesson.subject) && !item.formula) {
      throw new Error('Math key concept ' + (index + 1) + ' must include a formula.');
    }

    if (isMathSubject_(lesson.subject) && item.formula) {
      if (!/\\\[.*\\\]|\\\(.*\\\)/.test(item.formula)) {
        throw new Error('Math key concept ' + (index + 1) + ' formula must use KaTeX delimiters.');
      }
    }
  });

  lesson.practiceItems.forEach(function (item, index) {
    if (!item.question) throw new Error('Practice item ' + (index + 1) + ' has no question.');
    if (!item.acceptedAnswers || !item.acceptedAnswers.length) {
      throw new Error('Practice item ' + (index + 1) + ' has no accepted answer.');
    }
    if (item.type === 'multiple-choice' && item.options.length !== 4) {
      throw new Error('Practice item ' + (index + 1) + ' must have 4 options for multiple-choice.');
    }
    if (item.type === 'true-false' && item.options.length !== 2) {
      throw new Error('Practice item ' + (index + 1) + ' must have True and False options.');
    }
  });
}

/***************************************
 * 7. JSON EXTRACTION / REPAIR
 ***************************************/
function extractCandidateText_(apiResponse) {
  if (!apiResponse || !apiResponse.candidates || !apiResponse.candidates.length) return '';
  const parts = (((apiResponse.candidates[0] || {}).content || {}).parts || []);
  return parts.map(function (p) {
    return p && p.text ? p.text : '';
  }).join('').trim();
}

function extractApiError_(text) {
  try {
    const parsed = JSON.parse(text);
    return parsed && parsed.error && parsed.error.message
      ? parsed.error.message
      : '';
  } catch (err) {
    return '';
  }
}

function extractJsonText_(text) {
  let t = String(text || '').trim();
  t = t.replace(/^```json\s*/i, '').replace(/^```\s*/i, '').replace(/\s*```$/, '').trim();

  const firstBrace = t.indexOf('{');
  const lastBrace = t.lastIndexOf('}');
  if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
    t = t.substring(firstBrace, lastBrace + 1);
  }

  return t;
}

function repairLessonJson_(brokenJsonText, schema, apiKey, model) {
  const url =
    'https://generativelanguage.googleapis.com/v1beta/' +
    model +
    ':generateContent?key=' +
    encodeURIComponent(apiKey);

  const repairPrompt = [
    'Fix the following malformed JSON so that it becomes valid JSON.',
    'Return JSON only.',
    'Do not add commentary.',
    'Preserve the intended meaning.',
    'Ensure the output matches the provided schema exactly.',
    '',
    'Malformed JSON:',
    brokenJsonText
  ].join('\n');

  const payload = {
    contents: [
      {
        parts: [{ text: repairPrompt }]
      }
    ],
    generationConfig: {
      maxOutputTokens: 8192,
      responseMimeType: 'application/json',
      responseJsonSchema: schema
    }
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const text = response.getContentText();

  if (status !== 200) {
    throw new Error(extractApiError_(text) || ('Repair request failed with status ' + status + '.'));
  }

  const raw = JSON.parse(text);
  const outputText = extractCandidateText_(raw);
  if (!outputText) {
    throw new Error('Repair response was empty.');
  }

  return JSON.parse(extractJsonText_(outputText));
}

/***************************************
 * 8. HTML RENDERING
 ***************************************/
function renderLessonHtml_(lesson, options) {
  const opts = options || {};
  const forPrint = !!opts.forPrint;
  const copyMode = String(opts.copyMode || 'STUDENT').toUpperCase();
  const teacherView = copyMode === 'TEACHER';
  const layoutMode = String(opts.layoutMode || DEFAULT_LAYOUT_MODE).toUpperCase();
  const headerInfo = opts.headerInfo || {};
  const templateType = String(lesson.templateType || DEFAULT_TEMPLATE).toUpperCase();

  const topHeaderHtml = renderTopHeaderHtml_(lesson, teacherView, headerInfo);
  const lessonSheetClass = 'lesson-sheet layout-' + layoutMode.toLowerCase();
  const bodyHtml = renderTemplateBody_(lesson, {
    teacherView: teacherView,
    forPrint: forPrint,
    templateType: templateType
  });

  return `
    <div class="${lessonSheetClass}" data-copy-mode="${copyMode}" data-template="${escapeHtml_(templateType)}">
      ${topHeaderHtml}
      ${bodyHtml}
    </div>
  `;
}

function renderTopHeaderHtml_(lesson, teacherView, headerInfo) {
  return `
    <header class="lesson-header">
      <div class="doc-banner">
        ${headerInfo.schoolName ? `<div class="school-name editable" data-field="schoolName">${escapeHtml_(headerInfo.schoolName)}</div>` : ''}
        ${headerInfo.customHeader ? `<div class="custom-header editable" data-field="customHeader">${escapeHtml_(headerInfo.customHeader)}</div>` : ''}
      </div>
      <div class="eyebrow">${teacherView ? 'Teacher Copy' : 'Student Copy'}</div>
      <h1 class="lesson-title editable" data-field="title">${escapeHtml_(lesson.title)}</h1>
      <div class="lesson-meta">
        <div><span class="meta-label">Subject</span><span class="meta-value editable" data-field="subject">${escapeHtml_(lesson.subject)}</span></div>
        <div><span class="meta-label">Grade Level</span><span class="meta-value editable" data-field="gradeLevel">${escapeHtml_(lesson.gradeLevel)}</span></div>
        <div><span class="meta-label">Topic</span><span class="meta-value editable" data-field="topic">${escapeHtml_(lesson.topic)}</span></div>
        <div><span class="meta-label">Template</span><span class="meta-value">${escapeHtml_(prettyEnum_(lesson.templateType))}</span></div>
        <div><span class="meta-label">Difficulty</span><span class="meta-value">${escapeHtml_(prettyEnum_(lesson.difficulty))}</span></div>
        <div><span class="meta-label">Quarter</span><span class="meta-value editable" data-field="quarter">${escapeHtml_(headerInfo.quarter || '')}</span></div>
<div><span class="meta-label">Teacher Name</span><span class="meta-value editable" data-field="teacherName">${escapeHtml_(headerInfo.teacherName || '')}</span></div>
      </div>
    </header>
  `;
}

function renderTemplateBody_(lesson, options) {
  const opts = options || {};
  const teacherView = !!opts.teacherView;
  const forPrint = !!opts.forPrint;
  const templateType = String(opts.templateType || lesson.templateType || DEFAULT_TEMPLATE).toUpperCase();

  if (templateType === 'QUIZ') {
    return renderQuizTemplate_(lesson, teacherView, forPrint);
  }

  if (templateType === 'GUIDED_PRACTICE') {
    return renderGuidedPracticeTemplate_(lesson, teacherView, forPrint);
  }

  if (templateType === 'REVIEW') {
    return renderReviewTemplate_(lesson, teacherView, forPrint);
  }

  if (templateType === 'REMEDIATION') {
    return renderRemediationTemplate_(lesson, teacherView, forPrint);
  }

  if (templateType === 'ENRICHMENT') {
    return renderEnrichmentTemplate_(lesson, teacherView, forPrint);
  }

  return renderConceptTemplate_(lesson, teacherView, forPrint);
}

function renderConceptTemplate_(lesson, teacherView, forPrint) {
  return [
    renderObjectiveSection_(lesson),
    renderSuccessCriteriaSection_(lesson),
    renderVocabularySection_(lesson, false),
    renderSummarySection_(lesson, 'Lesson Overview'),
    renderFullConceptSection_(lesson),
    renderPracticeSection_(lesson, teacherView, forPrint, {
      title: 'Individual Practice',
      note: teacherView
        ? 'Teacher view shows accepted answers and hints.'
        : 'Answer all items. Use the interactive buttons for immediate checking.'
    }),
    renderTeacherNotesSection_(lesson, teacherView)
  ].join('');
}

function renderGuidedPracticeTemplate_(lesson, teacherView, forPrint) {
  return [
    renderObjectiveSection_(lesson),
    renderSuccessCriteriaSection_(lesson),
    renderVocabularySection_(lesson, true),
    renderGuidedPracticeFocusSection_(lesson),
    renderGuidedPracticeLearnSection_(lesson),
    renderGuidedPracticeModelSection_(lesson),
    renderPracticeSection_(lesson, teacherView, forPrint, {
      title: 'MASTER',
      note: teacherView
        ? 'Teacher copy shows accepted answers and hints after the Learn and Practice stages.'
        : 'Move through Learn and Practice first, then complete the Master tasks independently.'
    }),
    renderTeacherNotesSection_(lesson, teacherView)
  ].join('');
}

function renderGuidedPracticeFocusSection_(lesson) {
  return `
    <section>
      <h2 class="section-title">LEARN → PRACTICE → MASTER</h2>
      <div class="guided-focus-strip">
        <div class="guided-focus-title">Guided Learning Path</div>
        <div class="editable summary-text" data-field="lessonSummary">${escapeHtml_(lesson.lessonSummary)}</div>
      </div>
    </section>
  `;
}

function renderGuidedPracticeLearnSection_(lesson) {
  const conceptsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="card guided-learn-card" data-concept-index="${index}">
        <h3 class="concept-title editable" data-concept-field="heading">${escapeHtml_(item.heading)}</h3>

        <div class="guided-summary-box">
          <div class="concept-summary editable" data-concept-field="summary">${escapeHtml_(item.summary)}</div>
        </div>

        ${item.formula ? `
          <div class="mini-formula-box">
            <div class="mini-label">Useful Formula / Rule</div>
            <div class="formula-value formula-text" data-concept-field="formula">${escapeHtml_(item.formula)}</div>
          </div>
        ` : ''}

        <div class="guided-symbol-box">
          <div class="sub-card-label">Key Terms and Symbols</div>
          <div class="editable compact-note" data-concept-field="symbolMeaning">${escapeHtml_(item.symbolMeaning)}</div>
        </div>

        ${item.misconception ? `
          <div class="misconception-box">
            <strong>Watch Out:</strong>
            <div class="editable" data-concept-field="misconception">${escapeHtml_(item.misconception)}</div>
          </div>
        ` : ''}
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">LEARN</h2>
      <div class="section-note">Study the key ideas, formula or rule, important terms, and the common error to avoid.</div>
      <div class="guided-stage-grid">
        ${conceptsHtml}
      </div>
    </section>
  `;
}

function renderGuidedPracticeModelSection_(lesson) {
  const modelsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="guided-example-card" data-guided-model-index="${index}">
        <div class="guided-example-title">MODELED EXAMPLE</div>
        <h3 class="concept-title">${escapeHtml_(item.heading)}</h3>
        <div class="guided-transition-note">Follow the model carefully, then use the same process in the MASTER tasks.</div>
        <div class="example-box">
          <div>${escapeHtml_(item.workedExample)}</div>
        </div>
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">PRACTICE</h2>
      <div class="section-note">Use these guided examples to practice the process step by step before working independently.</div>
      ${modelsHtml}
    </section>
  `;
}

function renderReviewTemplate_(lesson, teacherView, forPrint) {
  return [
    renderObjectiveSection_(lesson),
    renderSuccessCriteriaSection_(lesson),
    renderVocabularySection_(lesson, true),
    renderReviewFocusSection_(lesson),
    renderReviewConceptSection_(lesson),
    renderPracticeSection_(lesson, teacherView, forPrint, {
      title: 'MIXED REVIEW',
      note: teacherView
        ? 'Teacher copy shows accepted answers and hints for quick checking.'
        : 'Use these items to review the key ideas and check your recall.'
    }),
    renderTeacherNotesSection_(lesson, teacherView)
  ].join('');
}

function renderRemediationTemplate_(lesson, teacherView, forPrint) {
  const splitIndex = Math.ceil((lesson.practiceItems || []).length / 2);

  return [
    renderObjectiveSection_(lesson),
    renderSuccessCriteriaSection_(lesson),
    renderVocabularySection_(lesson, true),
    renderRemediationFocusSection_(lesson),
    renderRemediationConceptSection_(lesson),
    renderRemediationModelSection_(lesson),
    renderPracticeGroupSection_(lesson, teacherView, forPrint, {
      title: 'TRY WITH HELP',
      note: teacherView
        ? 'Use these first items as supported practice. Teacher view shows accepted answers and hints.'
        : 'Start with these supported items. Use the hints when needed and focus on one step at a time.',
      className: 'try-help-section'
    }, 0, splitIndex),
    renderPracticeGroupSection_(lesson, teacherView, forPrint, {
      title: 'CHECK YOURSELF',
      note: teacherView
        ? 'Use these items to see whether students can apply the skill more independently.'
        : 'Complete these items on your own to check your understanding.',
      className: 'check-yourself-section'
    }, splitIndex, lesson.practiceItems.length),
    renderTeacherNotesSection_(lesson, teacherView)
  ].join('');
}

function renderEnrichmentTemplate_(lesson, teacherView, forPrint) {
  const splitIndex = Math.ceil((lesson.practiceItems || []).length / 2);

  return [
    renderObjectiveSection_(lesson),
    renderSuccessCriteriaSection_(lesson),
    renderVocabularySection_(lesson, true),
    renderEnrichmentRecallSection_(lesson),
    renderEnrichmentConceptSection_(lesson),
    renderPracticeGroupSection_(lesson, teacherView, forPrint, {
      title: 'CHALLENGE',
      note: teacherView
        ? 'These items push beyond routine practice and may need stronger reasoning.'
        : 'Tackle these challenge items by applying the ideas carefully and explaining your thinking.',
      className: 'challenge-section'
    }, 0, splitIndex),
    renderPracticeGroupSection_(lesson, teacherView, forPrint, {
      title: 'TRANSFER',
      note: teacherView
        ? 'These items check whether students can transfer the skill to less familiar situations.'
        : 'Apply the ideas in new or less familiar situations and justify your thinking clearly.',
      className: 'transfer-section'
    }, splitIndex, lesson.practiceItems.length),
    renderTeacherNotesSection_(lesson, teacherView)
  ].join('');
}

function renderRemediationFocusSection_(lesson) {
  return `
    <section>
      <h2 class="section-title">UNDERSTAND → FOLLOW → TRY WITH HELP → CHECK YOURSELF</h2>
      <div class="remediation-strip">
        <div class="remediation-strip-title">Support Path</div>
        <div class="editable summary-text" data-field="lessonSummary">${escapeHtml_(lesson.lessonSummary)}</div>
      </div>
    </section>
  `;
}

function renderRemediationConceptSection_(lesson) {
  const conceptsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="remediation-card" data-concept-index="${index}">
        <div class="remediation-model-title">UNDERSTAND</div>
        <h3 class="concept-title editable" data-concept-field="heading">${escapeHtml_(item.heading)}</h3>

        <div class="concept-summary editable" data-concept-field="summary">${escapeHtml_(item.summary)}</div>

        ${item.formula ? `
          <div class="mini-formula-box">
            <div class="mini-label">Rule / Formula</div>
            <div class="formula-value formula-text" data-concept-field="formula">${escapeHtml_(item.formula)}</div>
          </div>
        ` : ''}

        <div class="remediation-help-box">
          <div class="sub-card-label">Helpful Terms and Symbols</div>
          <div class="editable compact-note" data-concept-field="symbolMeaning">${escapeHtml_(item.symbolMeaning)}</div>
        </div>

        ${item.misconception ? `
          <div class="misconception-box">
            <strong>Watch Out:</strong>
            <div class="editable" data-concept-field="misconception">${escapeHtml_(item.misconception)}</div>
          </div>
        ` : ''}
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">UNDERSTAND</h2>
      <div class="section-note">Study the idea first. Focus on the rule, the key terms, and the common mistake to avoid.</div>
      <div class="remediation-grid">
        ${conceptsHtml}
      </div>
    </section>
  `;
}

function renderRemediationModelSection_(lesson) {
  const modelsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="remediation-model-card" data-remediation-model-index="${index}">
        <div class="remediation-model-title">FOLLOW THE EXAMPLE</div>
        <h3 class="concept-title">${escapeHtml_(item.heading)}</h3>
        <div class="guided-transition-note">Read the worked example carefully, then copy the same steps in the supported practice.</div>
        <div class="example-box">
          <div>${escapeHtml_(item.workedExample)}</div>
        </div>
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">FOLLOW</h2>
      <div class="section-note">Use these worked examples as step-by-step guides before answering on your own.</div>
      ${modelsHtml}
    </section>
  `;
}

function renderEnrichmentRecallSection_(lesson) {
  return `
    <section>
      <h2 class="section-title">RECALL → THINK DEEPER → CHALLENGE → TRANSFER</h2>
      <div class="enrichment-strip">
        <div class="enrichment-strip-title">Extension Path</div>
        <div class="editable summary-text" data-field="lessonSummary">${escapeHtml_(lesson.lessonSummary)}</div>
      </div>
    </section>
  `;
}

function renderEnrichmentConceptSection_(lesson) {
  const conceptsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="enrichment-card" data-concept-index="${index}">
        <div class="enrichment-card-title">THINK DEEPER</div>
        <h3 class="concept-title editable" data-concept-field="heading">${escapeHtml_(item.heading)}</h3>

        <div class="concept-summary editable" data-concept-field="summary">${escapeHtml_(item.summary)}</div>

        ${item.formula ? `
          <div class="review-rule-box">
            <div class="mini-label">Key Rule / Relationship</div>
            <div class="formula-value formula-text" data-concept-field="formula">${escapeHtml_(item.formula)}</div>
          </div>
        ` : ''}

        <div class="enrichment-note-box">
          <div class="sub-card-label">Big Idea to Keep in Mind</div>
          <div class="editable compact-note" data-concept-field="symbolMeaning">${escapeHtml_(item.symbolMeaning)}</div>
        </div>

        ${item.misconception ? `
          <div class="misconception-box">
            <strong>Avoid This Shortcut Error:</strong>
            <div class="editable" data-concept-field="misconception">${escapeHtml_(item.misconception)}</div>
          </div>
        ` : ''}
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">THINK DEEPER</h2>
      <div class="section-note">Review the big ideas first, then use them in more demanding and less familiar tasks.</div>
      <div class="enrichment-grid">
        ${conceptsHtml}
      </div>
    </section>
  `;
}

function renderPracticeGroupSection_(lesson, teacherView, forPrint, config, startIndex, endIndex) {
  const cfg = config || {};
  const title = cfg.title || 'Practice';
  const note = cfg.note || '';
  const className = cfg.className || '';
  const subset = (lesson.practiceItems || []).slice(startIndex, endIndex);

  if (!subset.length) return '';

  const practiceHtml = subset.map(function (item, offset) {
    return renderPracticeItemHtml_(item, startIndex + offset, teacherView, forPrint);
  }).join('');

  return `
    <section class="practice-section ${className}">
      <h2 class="section-title">${escapeHtml_(title)}</h2>
      ${note ? `<div class="section-note">${escapeHtml_(note)}</div>` : ''}
      ${practiceHtml}
    </section>
  `;
}

function renderQuizTemplate_(lesson, teacherView, forPrint) {
  return [
    teacherView ? renderQuizTeacherOverviewSection_(lesson) : '',
    renderQuizInfoSection_(teacherView),
    renderQuizDirectionsSection_(teacherView, forPrint),
    renderPracticeSection_(lesson, teacherView, forPrint, {
      title: 'Quiz Items',
      note: teacherView
        ? 'Teacher copy includes accepted answers and hints beside each item.'
        : 'Complete all items independently and review your answers before submitting.'
    }),
    renderTeacherNotesSection_(lesson, teacherView)
  ].join('');
}

function renderObjectiveSection_(lesson) {
  return `
    <section>
      <h2 class="section-title">Objective</h2>
      <div class="card">
        <div class="editable objective-text" data-field="objective">${escapeHtml_(lesson.objective)}</div>
      </div>
    </section>
  `;
}

function renderSuccessCriteriaSection_(lesson) {
  if (!lesson.successCriteria || !lesson.successCriteria.length) return '';

  return `
    <section>
      <h2 class="section-title">Success Criteria</h2>
      <div class="card">
        <ul class="criteria-list">
          ${lesson.successCriteria.map(function (item, index) {
    return `<li class="editable" data-criteria-index="${index}">${escapeHtml_(item)}</li>`;
  }).join('')}
        </ul>
      </div>
    </section>
  `;
}

function renderVocabularySection_(lesson, compact) {
  if (!lesson.vocabulary || !lesson.vocabulary.length) return '';

  if (compact) {
    return `
      <section>
        <h2 class="section-title">Vocabulary</h2>
        <div class="card">
          <ul class="compact-vocab-list">
            ${lesson.vocabulary.map(function (item, index) {
      return `
                <li class="vocab-item" data-vocab-index="${index}">
                  <span class="compact-vocab-term vocab-term editable" data-vocab-field="term">${escapeHtml_(item.term)}</span>
                  <span class="editable vocab-definition" data-vocab-field="definition">: ${escapeHtml_(item.definition)}</span>
                </li>
              `;
    }).join('')}
          </ul>
        </div>
      </section>
    `;
  }

  return `
    <section>
      <h2 class="section-title">Vocabulary</h2>
      <div class="card">
        <div class="vocab-grid">
          ${lesson.vocabulary.map(function (item, index) {
    return `
              <div class="vocab-item" data-vocab-index="${index}">
                <div class="vocab-term editable" data-vocab-field="term">${escapeHtml_(item.term)}</div>
                <div class="vocab-definition editable" data-vocab-field="definition">${escapeHtml_(item.definition)}</div>
              </div>
            `;
  }).join('')}
        </div>
      </div>
    </section>
  `;
}

function renderSummarySection_(lesson, title) {
  if (!lesson.lessonSummary) return '';

  return `
    <section>
      <h2 class="section-title">${escapeHtml_(title || 'Lesson Overview')}</h2>
      <div class="card">
        <div class="section-note editable summary-text" data-field="lessonSummary">${escapeHtml_(lesson.lessonSummary)}</div>
      </div>
    </section>
  `;
}

function renderFullConceptSection_(lesson) {
  const conceptsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="card concept-card" data-concept-index="${index}">
        <h3 class="concept-title editable" data-concept-field="heading">${escapeHtml_(item.heading)}</h3>
        <div class="concept-summary editable" data-concept-field="summary">${escapeHtml_(item.summary)}</div>
        ${item.formula ? `
          <div class="formula-box">
            <div class="formula-label">Formula</div>
            <div class="formula-value formula-text" data-concept-field="formula">${escapeHtml_(item.formula)}</div>
          </div>
        ` : ''}
        <div class="sub-card">
          <div class="sub-card-label">Meaning of Symbols / Important Terms</div>
          <div class="editable" data-concept-field="symbolMeaning">${escapeHtml_(item.symbolMeaning)}</div>
        </div>
        <div class="example-box">
          <strong>Worked Example:</strong>
          <div class="editable" data-concept-field="workedExample">${escapeHtml_(item.workedExample)}</div>
        </div>
        ${item.misconception ? `
          <div class="misconception-box">
            <strong>Common Misconception:</strong>
            <div class="editable" data-concept-field="misconception">${escapeHtml_(item.misconception)}</div>
          </div>
        ` : ''}
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">Key Concepts</h2>
      <div class="section-note">Use these explanations and examples for direct instruction.</div>
      ${conceptsHtml}
    </section>
  `;
}

function renderGuidedPracticeConceptSection_(lesson) {
  const conceptsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="card concept-card guided-practice-card" data-concept-index="${index}">
        <div class="practice-number">Concept ${index + 1}</div>
        <h3 class="concept-title editable" data-concept-field="heading">${escapeHtml_(item.heading)}</h3>

        <div class="sub-card">
          <div class="sub-card-label">Teach It</div>
          <div class="concept-summary editable" data-concept-field="summary">${escapeHtml_(item.summary)}</div>
        </div>

        ${item.formula ? `
          <div class="formula-box">
            <div class="formula-label">Useful Formula / Rule</div>
            <div class="formula-value formula-text" data-concept-field="formula">${escapeHtml_(item.formula)}</div>
          </div>
        ` : ''}

        <div class="sub-card">
          <div class="sub-card-label">Key Terms and Symbols</div>
          <div class="editable compact-note" data-concept-field="symbolMeaning">${escapeHtml_(item.symbolMeaning)}</div>
        </div>

        <div class="example-box">
          <strong>Try It Together:</strong>
          <div class="editable" data-concept-field="workedExample">${escapeHtml_(item.workedExample)}</div>
        </div>

        ${item.misconception ? `
          <div class="misconception-box">
            <strong>Watch Out:</strong>
            <div class="editable" data-concept-field="misconception">${escapeHtml_(item.misconception)}</div>
          </div>
        ` : ''}
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">Guided Learning</h2>
      <div class="section-note">Move from explanation to supported example before asking students to work independently.</div>
      ${conceptsHtml}
    </section>
  `;
}

function renderReviewFocusSection_(lesson) {
  return `
    <section>
      <h2 class="section-title">REVIEW AT A GLANCE</h2>
      <div class="review-focus-strip">
        <div class="review-focus-title">Quick Focus</div>
        <div class="editable summary-text" data-field="lessonSummary">${escapeHtml_(lesson.lessonSummary)}</div>
      </div>
    </section>
  `;
}

function renderReviewConceptSection_(lesson) {
  const conceptsHtml = lesson.keyConcepts.map(function (item, index) {
    return `
      <article class="review-mini-card quick-review-card" data-concept-index="${index}">
        <div class="review-mini-title">QUICK REMINDER</div>
        <h3 class="concept-title editable" data-concept-field="heading">${escapeHtml_(item.heading)}</h3>

        <div class="concept-summary editable" data-concept-field="summary">${escapeHtml_(item.summary)}</div>

        ${item.formula ? `
          <div class="review-rule-box">
            <div class="mini-label">Formula / Rule</div>
            <div class="formula-value formula-text" data-concept-field="formula">${escapeHtml_(item.formula)}</div>
          </div>
        ` : ''}

        <div class="review-reminder-box">
          <div class="sub-card-label">Remember This</div>
          <div class="editable compact-note" data-concept-field="symbolMeaning">${escapeHtml_(item.symbolMeaning)}</div>
        </div>

        ${item.misconception ? `
          <div class="misconception-box">
            <strong>Common Error:</strong>
            <div class="editable" data-concept-field="misconception">${escapeHtml_(item.misconception)}</div>
          </div>
        ` : ''}
      </article>
    `;
  }).join('');

  return `
    <section>
      <h2 class="section-title">QUICK REMINDERS</h2>
      <div class="section-note">Scan these reminders first, then answer the review items.</div>
      <div class="review-list">
        ${conceptsHtml}
      </div>
    </section>
  `;
}

function renderQuizTeacherOverviewSection_(lesson) {
  return `
    <section>
      <h2 class="section-title">Assessment Focus</h2>
      <div class="quiz-focus-card">
        <div class="quiz-focus-title">Objective</div>
        <div class="editable quiz-focus-text objective-text" data-field="objective">${escapeHtml_(lesson.objective)}</div>

        ${lesson.successCriteria && lesson.successCriteria.length ? `
          <div class="quiz-focus-block">
            <div class="quiz-focus-title">What This Checks</div>
            <ul class="criteria-list">
              ${lesson.successCriteria.map(function (item, index) {
    return `<li class="editable" data-criteria-index="${index}">${escapeHtml_(item)}</li>`;
  }).join('')}
            </ul>
          </div>
        ` : ''}
      </div>
    </section>
  `;
}

function renderQuizInfoSection_(teacherView) {
  return `
    <section>
      <h2 class="section-title">${teacherView ? 'Assessment Record' : 'Student Information'}</h2>
      <div class="quiz-info-card">
        <div class="quiz-meta-lines">
          <div class="quiz-line">
            <span class="quiz-line-label">Name</span>
            <span class="quiz-line-fill"></span>
          </div>
          <div class="quiz-line">
            <span class="quiz-line-label">Date</span>
            <span class="quiz-line-fill"></span>
          </div>
          <div class="quiz-line">
            <span class="quiz-line-label">Score</span>
            <span class="quiz-line-fill"></span>
          </div>
        </div>
      </div>
    </section>
  `;
}

function renderQuizDirectionsSection_(teacherView, forPrint) {
  const previewOnlyNote = (!teacherView && !forPrint)
    ? '<li>Preview buttons are for on-screen checking only and will not appear in print.</li>'
    : '';

  return `
    <section>
      <h2 class="section-title">Directions</h2>
      <div class="directions-box">
        ${teacherView
      ? 'Use this copy as the checking and reference version of the assessment.'
      : 'Read each item carefully and answer the questions independently.'}
        <ul class="directions-list">
          <li>Answer every question.</li>
          <li>Work neatly and review your answers before submitting.</li>
          ${previewOnlyNote}
        </ul>
      </div>
    </section>
  `;
}

function renderPracticeSection_(lesson, teacherView, forPrint, config) {
  const cfg = config || {};
  const title = cfg.title || 'Individual Practice';
  const note = cfg.note || (teacherView
    ? 'Teacher view shows accepted answers and hints.'
    : 'Answer all items. Use the interactive buttons for immediate checking.');

  const practiceHtml = lesson.practiceItems.map(function (item, index) {
    return renderPracticeItemHtml_(item, index, teacherView, forPrint);
  }).join('');

  return `
    <section class="practice-section">
      <h2 class="section-title">${escapeHtml_(title)}</h2>
      <div class="section-note">${escapeHtml_(note)}</div>
      ${practiceHtml}
    </section>
  `;
}

function renderTeacherNotesSection_(lesson, teacherView) {
  if (!teacherView || !lesson.teacherNotes || !lesson.teacherNotes.length) return '';

  return `
    <section>
      <h2 class="section-title">Teacher Notes</h2>
      <div class="card">
        <ul class="teacher-note-list">
          ${lesson.teacherNotes.map(function (note, index) {
    return `<li class="editable" data-note-index="${index}">${escapeHtml_(note)}</li>`;
  }).join('')}
        </ul>
      </div>
    </section>
  `;
}

function renderPracticeItemHtml_(item, index, teacherView, forPrint) {
  const answerText = item.acceptedAnswers.join(' | ');
  let controlsHtml = '';

  if (item.type === 'multiple-choice') {
    const optionMarkup = item.options.map(function (opt, i) {
      const letter = String.fromCharCode(65 + i);
      return `
        <label class="mc-choice">
          <input type="radio" name="practice_${index}" value="${escapeHtml_(opt)}">
          <span class="mc-letter">${letter}.</span>
          <span class="mc-choice-text">${escapeHtml_(opt)}</span>
        </label>
      `;
    }).join('');

    controlsHtml = forPrint ? `
      <div class="option-preview">
        ${item.options.map(function (opt, i) {
      return `<div class="option-line"><strong>${String.fromCharCode(65 + i)}.</strong> ${escapeHtml_(opt)}</div>`;
    }).join('')}
      </div>
      <div class="print-answer-space"></div>
    ` : `
      <div class="practice-controls practice-controls-block">
        <div class="mc-choice-list">
          ${optionMarkup}
        </div>
        <div class="mc-action-row">
          <button class="btn btn-primary" onclick="checkAnswer(${index})">Check</button>
          <button class="btn btn-light" onclick="showAnswer(${index})">Show Answer</button>
        </div>
      </div>
    `;
  } else if (item.type === 'true-false') {
    controlsHtml = forPrint ? '<div class="print-answer-space"></div>' : `
      <div class="practice-controls">
        <select id="practice_${index}" class="answer-input">
          <option value="">Select an answer</option>
          <option value="True">True</option>
          <option value="False">False</option>
        </select>
        <button class="btn btn-primary" onclick="checkAnswer(${index})">Check</button>
        <button class="btn btn-light" onclick="showAnswer(${index})">Show Answer</button>
      </div>
    `;
  } else {
    controlsHtml = forPrint ? '<div class="print-answer-space lines-2"></div>' : `
      <div class="practice-controls">
        <input id="practice_${index}" class="answer-input" type="text" autocomplete="off" placeholder="Type your answer">
        <button class="btn btn-primary" onclick="checkAnswer(${index})">Check</button>
        <button class="btn btn-light" onclick="showAnswer(${index})">Show Answer</button>
      </div>
    `;
  }

  return `
    <article class="card practice-card" data-practice-index="${index}">
      <div class="practice-number">Item ${index + 1} • ${escapeHtml_(item.type)}</div>
      <div class="practice-question editable" data-practice-field="question">${escapeHtml_(item.question)}</div>
      ${controlsHtml}
      ${forPrint ? '' : `<div id="feedback_${index}" class="feedback"></div>`}
      ${teacherView
      ? `<div class="teacher-answer-box">
            <strong>Accepted answer(s):</strong> <span class="editable" data-practice-field="answers">${escapeHtml_(answerText)}</span>
            ${item.hint ? `<div class="hint-line"><strong>Hint:</strong> <span class="editable" data-practice-field="hint">${escapeHtml_(item.hint)}</span></div>` : ''}
          </div>`
      : `<div id="answer_${index}" class="answer-reveal hidden">
            <strong>Accepted answer(s):</strong> ${escapeHtml_(answerText)}
            ${item.hint ? `<div class="hint-line"><strong>Hint:</strong> ${escapeHtml_(item.hint)}</div>` : ''}
          </div>`
    }
    </article>
  `;
}

/***************************************
 * 9. PRINT DOCUMENT BUILD
 ***************************************/
function buildPrintableDocument_(bodyHtml, options) {
  const opts = options || {};
  const layoutMode = opts.layoutMode || DEFAULT_LAYOUT_MODE;
  const footerText = opts.footerText || APP_TITLE;
  const documentTitle = opts.documentTitle || APP_TITLE;

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>${escapeHtml_(documentTitle)}</title>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css">
        <style>
          ${getSharedStyles_(layoutMode)}
          ${getPrintStyles_()}
        </style>
      </head>
      <body class="print-mode">
        ${bodyHtml}
        <div class="pdf-footnote">${escapeHtml_(footerText)}</div>
        <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js"></script>
        <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/contrib/auto-render.min.js"></script>
        <script>
  function renderMathForPrint() {
    if (typeof renderMathInElement !== 'function') {
      window.setTimeout(renderMathForPrint, 150);
      return;
    }

    renderMathInElement(document.body, {
      delimiters: [
        { left: '\\\\[', right: '\\\\]', display: true },
        { left: '\\\\(', right: '\\\\)', display: false }
      ],
      throwOnError: false
    });

    Promise.resolve(document.fonts ? document.fonts.ready : null).then(function () {
      window.requestAnimationFrame(function () {
        window.requestAnimationFrame(function () {
          window.print();
        });
      });
    });
  }

  window.addEventListener('load', function () {
    window.setTimeout(renderMathForPrint, 150);
  });
</script>
      </body>
    </html>
  `;
}

function getSharedStyles_(layoutMode) {
  return [
    '*{box-sizing:border-box;}',
    'body{margin:0;font-family:Arial,Helvetica,sans-serif;background:#f5f7fb;color:#172b4d;}',
    getLessonStyles_(),
    '.btn{border:none;border-radius:10px;padding:11px 16px;font-size:13px;font-weight:700;cursor:pointer;}',
    '.btn-primary{background:#0052cc;color:#fff;}',
    '.btn-light{background:#eef3fb;color:#223b53;}',
    '.feedback.success{color:#0c6b2f;}',
    '.feedback.error{color:#b42318;}',
    '.pdf-footnote{display:none;}'
  ].join('');
}

function getPrintStyles_() {
  return [
    '@page{size:Letter; margin:0.5in;}',
    'html,body{width:100%;}',
    'body.print-mode{background:#fff !important;color:#000;}',
    'body.print-mode .lesson-sheet{width:100%;max-width:none;padding:0;margin:0 auto;}',
    'body.print-mode .lesson-header, body.print-mode .card{box-shadow:none;}',
    'body.print-mode .lesson-header{padding:14px;margin-bottom:10px;}',
    'body.print-mode .card{padding:12px;margin-bottom:10px;}',
    'body.print-mode .lesson-title{font-size:22px;line-height:1.15;}',
    'body.print-mode .section-title{font-size:15px;margin:14px 0 4px;break-after:avoid;page-break-after:avoid;}',
    'body.print-mode section{break-inside:auto;page-break-inside:auto;}',
    'body.print-mode .section-note{font-size:12px;line-height:1.45;margin-bottom:8px;}',
    'body.print-mode .meta-value{font-size:12px;}',
    'body.print-mode .concept-summary, body.print-mode .practice-question, body.print-mode .vocab-definition, body.print-mode .teacher-answer-box, body.print-mode .answer-reveal, body.print-mode .sub-card, body.print-mode .example-box, body.print-mode .misconception-box, body.print-mode .objective-text, body.print-mode .summary-text{font-size:12px;line-height:1.45;white-space:pre-line;overflow-wrap:anywhere;word-break:break-word;}',
    'body.print-mode .formula-value{font-size:13px;line-height:1.4;}',
    'body.print-mode .lesson-meta{display:grid !important;grid-template-columns:repeat(3,minmax(0,1fr)) !important;gap:6px !important;}',
    'body.print-mode .lesson-meta>div{padding:8px 10px;border-radius:10px;}',
    'body.print-mode .meta-label{font-size:9px;line-height:1.1;margin-bottom:2px;}',
    'body.print-mode .meta-value{font-size:11px;line-height:1.2;}',
    'body.print-mode .vocab-grid{grid-template-columns:repeat(auto-fit,minmax(2in,1fr));gap:8px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .section-title{margin:10px 0 4px;font-size:14px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .section-note{font-size:11px;line-height:1.35;margin-bottom:6px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .review-focus-strip{padding:10px 12px;margin-bottom:10px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .review-focus-title{font-size:10px;margin-bottom:4px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .compact-vocab-list{columns:2;column-gap:18px;padding-left:16px;line-height:1.45;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .compact-vocab-list li{margin-bottom:6px;break-inside:avoid;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .review-list{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .review-mini-card, body.print-mode .lesson-sheet[data-template="REVIEW"] .quick-review-card{padding:10px 12px;margin-bottom:0;break-inside:avoid;page-break-inside:avoid;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .review-mini-title{font-size:10px;margin-bottom:6px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .concept-title{font-size:14px;line-height:1.25;margin:0 0 6px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .concept-summary{font-size:11px;line-height:1.4;margin-bottom:6px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .review-rule-box{padding:8px 10px;margin:6px 0 8px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .review-reminder-box{padding:8px 10px;margin-top:6px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .misconception-box{padding:8px 10px;margin-top:6px;font-size:11px;line-height:1.35;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .practice-card{padding:10px 12px;margin-bottom:8px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .practice-number{font-size:10px;margin-bottom:6px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .practice-question{font-size:12px;line-height:1.45;margin-bottom:8px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .option-preview{padding:8px 10px;margin:8px 0 10px;}',
    'body.print-mode .lesson-sheet[data-template="REVIEW"] .print-answer-space{margin-top:12px;}',
    'body.print-mode .btn, body.print-mode .feedback, body.print-mode .answer-reveal.hidden{display:none !important;}',
    'body.print-mode .practice-controls{display:block;}',
    'body.print-mode .answer-input{display:block;width:100%;border:1px solid #999;min-width:0;}',
    '@media print and (max-width: 700px){body.print-mode .lesson-sheet[data-template="REVIEW"] .review-list{grid-template-columns:1fr;}body.print-mode .lesson-sheet[data-template="REVIEW"] .compact-vocab-list{columns:1;}}',
    'body.print-mode .btn, body.print-mode .feedback, body.print-mode .answer-reveal.hidden{display:none !important;}',
    'body.print-mode .concept-card, body.print-mode .practice-card{break-inside:auto;page-break-inside:auto;}',
    'body.print-mode .formula-box, body.print-mode .sub-card, body.print-mode .example-box, body.print-mode .misconception-box{break-inside:avoid;page-break-inside:avoid;}',
    'body.print-mode .editable.active-edit{outline:none;background:transparent;}',
    'body.print-mode .pdf-footnote{display:block !important;margin-top:0.35in;padding-top:10px;text-align:center;font-size:10px;line-height:1.2;color:#666;border-top:1px solid #ddd;break-inside:avoid;page-break-inside:avoid;}'
  ].join('');
}

/***************************************
 * 10. HELPERS
 ***************************************/
function normalizeDisplayText_(value) {
  return String(value == null ? '' : value)
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\s*\\\[\s*/g, '\n\\[')
    .replace(/\s*\\\]\s*/g, '\\]\n')
    .replace(/([^\n])\s+(Problem\s*:)/gi, '$1\n$2')
    .replace(/([^\n])\s+(Step\s+\d+\s*:)/gi, '$1\n$2')
    .replace(/([^\n])\s+(Answer\s*:)/gi, '$1\n$2')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function slugify_(text) {
  return String(text || 'lesson-guide')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .replace(/-{2,}/g, '-');
}

function toBoolean_(value, defaultValue) {
  if (value === true || value === false) return value;
  if (typeof value === 'string') {
    const v = value.toLowerCase().trim();
    if (v === 'true' || v === 'yes' || v === '1' || v === 'on') return true;
    if (v === 'false' || v === 'no' || v === '0' || v === 'off') return false;
  }
  return !!defaultValue;
}

function toBoundedInteger_(value, fallback, min, max) {
  const num = parseInt(value, 10);
  if (isNaN(num)) return fallback;
  return Math.max(min, Math.min(max, num));
}

function isMathSubject_(subject) {
  return String(subject || '').toLowerCase().indexOf('math') !== -1;
}

function prettyEnum_(value) {
  return String(value || '')
    .toLowerCase()
    .split('_')
    .map(function (part) {
      return part ? part.charAt(0).toUpperCase() + part.slice(1) : '';
    })
    .join(' ');
}

function getLayoutSpacing_(layoutMode) {
  const mode = String(layoutMode || DEFAULT_LAYOUT_MODE).toUpperCase();

  if (mode === 'COMPACT') {
    return {
      sheetPadding: '18px',
      blockPadding: '14px',
      blockGap: '10px'
    };
  }

  if (mode === 'SPACIOUS') {
    return {
      sheetPadding: '30px',
      blockPadding: '22px',
      blockGap: '18px'
    };
  }

  return {
    sheetPadding: '24px',
    blockPadding: '18px',
    blockGap: '14px'
  };
}

function getHeaderInfoFromRequest_(request) {
  const req = request || {};
  return {
    schoolName: String(req.schoolName || '').trim(),
    teacherName: String(req.teacherName || '').trim(),
    quarter: String(req.quarter || '').trim(),
    customHeader: String(req.customHeader || '').trim()
  };
}


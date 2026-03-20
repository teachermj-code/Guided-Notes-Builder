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
    '.lesson-sheet{max-width:900px;margin:0 auto;padding:24px;}',
    '.lesson-sheet.layout-compact{padding:18px;}',
    '.lesson-sheet.layout-spacious{padding:30px;}',
    '.lesson-header{background:#ffffff;border:1px solid #dfe3eb;border-radius:18px;padding:18px;margin-bottom:14px;}',
    '.lesson-sheet.layout-compact .lesson-header{padding:14px;margin-bottom:10px;}',
    '.lesson-sheet.layout-spacious .lesson-header{padding:22px;margin-bottom:18px;}',
    '.doc-banner{margin-bottom:18px;}',
    '.school-name{font-size:15px;font-weight:700;color:#0f2747;margin-bottom:4px;}',
    '.custom-header{font-size:12px;color:#5b6b82;margin-bottom:8px;}',
    '.eyebrow{font-size:11px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#0052cc;margin-bottom:8px;}',
    '.lesson-title{margin:0;font-size:28px;line-height:1.2;color:#0f2747;}',
    '.lesson-meta{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin-top:18px;}',
    '.lesson-meta>div{background:#f8fbff;border:1px solid #d9e6ff;border-radius:12px;padding:12px;}',
    '.meta-label{display:block;font-size:11px;font-weight:700;text-transform:uppercase;color:#5b6b82;margin-bottom:4px;}',
    '.meta-value{display:block;font-size:14px;font-weight:600;color:#1c2e45;}',
    '.section-title{font-size:18px;margin:24px 0 6px;color:#0052cc;}',
    '.section-note{font-size:13px;color:#5b6b82;margin-bottom:14px;line-height:1.6;}',
    '.card{background:#ffffff;border:1px solid #dfe3eb;border-radius:16px;padding:18px;margin-bottom:14px;}',
    '.lesson-sheet.layout-compact .card{padding:14px;margin-bottom:10px;}',
    '.lesson-sheet.layout-spacious .card{padding:22px;margin-bottom:18px;}',
    '.concept-card,.practice-card{break-inside:avoid;page-break-inside:avoid;}',
    '.concept-title{margin:0 0 8px;font-size:16px;color:#102a43;}',
    '.concept-summary{margin:0 0 12px;font-size:14px;line-height:1.7;}',
    '.formula-box{background:#f5f9ff;border:1px solid #d9e6ff;border-radius:12px;padding:12px;margin-bottom:12px;}',
    '.formula-label{font-size:11px;font-weight:700;text-transform:uppercase;color:#5b6b82;margin-bottom:6px;}',
    '.formula-value{font-size:15px;line-height:1.7;}',
    '.sub-card{background:#fafcff;border:1px solid #e5ebf5;border-radius:12px;padding:12px;margin-bottom:12px;font-size:13px;line-height:1.6;}',
    '.sub-card-label{font-size:11px;font-weight:700;text-transform:uppercase;color:#5b6b82;margin-bottom:6px;}',
    '.example-box{background:#f3f8ff;border:1px solid #d9e6ff;border-radius:12px;padding:12px;font-size:13px;line-height:1.7;margin-bottom:10px;}',
    '.misconception-box{background:#fff8e6;border:1px solid #f2d27c;border-radius:12px;padding:12px;font-size:13px;line-height:1.6;}',
    '.criteria-list,.teacher-note-list{margin:0;padding-left:20px;line-height:1.7;}',
    '.vocab-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;}',
    '.vocab-item{background:#fafcff;border:1px solid #e5ebf5;border-radius:12px;padding:12px;}',
    '.vocab-term{font-weight:700;color:#102a43;margin-bottom:6px;}',
    '.vocab-definition{font-size:13px;line-height:1.6;}',
    '.practice-number{font-size:12px;font-weight:700;color:#5b6b82;text-transform:uppercase;margin-bottom:8px;}',
    '.practice-question{font-size:14px;line-height:1.7;margin-bottom:12px;}',
    '.practice-controls{display:flex;flex-wrap:wrap;gap:10px;align-items:center;}',
    '.practice-controls-block{display:block;width:100%;}',
    '.mc-choice-list{display:flex;flex-direction:column;gap:10px;width:100%;}',
    '.mc-choice{display:flex;align-items:flex-start;gap:10px;padding:10px 12px;border:1px solid #dfe3eb;border-radius:12px;background:#fafcff;}',
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
    '.option-preview{margin:10px 0 12px;padding:10px 12px;background:#fafcff;border:1px solid #e5ebf5;border-radius:12px;}',
    '.option-line{margin:4px 0;font-size:13px;line-height:1.5;}',
    '.print-answer-space{height:28px;border-bottom:1px solid #7a869a;margin-top:18px;}',
    '.print-answer-space.lines-2{height:56px;}',
    '.editable.active-edit{outline:2px dashed #7aa7ff;outline-offset:2px;background:#f8fbff;cursor:text;}',
    '@media (max-width: 760px){.lesson-meta,.vocab-grid{grid-template-columns:1fr;}}'
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

  return [
    'You are an expert instructional content generator for classroom teachers.',
    'Create a classroom-ready lesson guide in JSON only.',
    '',
    'Required output fields:',
    '1. title',
    '2. subject',
    '3. gradeLevel',
    '4. topic',
    '5. templateType',
    '6. difficulty',
    '7. objective',
    '8. successCriteria array',
    '9. vocabulary array of objects with term and definition',
    '10. lessonSummary',
    '11. keyConcepts array',
    '12. practiceItems array',
    '13. teacherNotes array',
    '',
    'Key concept rules:',
'- Each key concept must include: heading, summary, formula, symbolMeaning, workedExample, misconception.',
'- For non-math topics, formula may be an empty string only if no formula is appropriate.',
'- symbolMeaning must explain symbols or important vocabulary used in the formula or concept.',
'- workedExample must begin with a short problem or context before the steps.',
'- For Mathematics, present workedExample in this order: Problem, Step 1, Step 2, Step 3, Answer.',
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
    '- Use a realistic mix of item types unless the template strongly suggests otherwise.',
    '- For short-answer, options must be an empty array.',
    '- For multiple-choice, provide exactly 4 options.',
    '- For true-false, options must be ["True","False"].',
    '- acceptedAnswers must contain at least one expected answer string.',
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
    'Subject-specific guidance:',
    subjectGuidance,
    '',
    'Template-specific guidance:',
    templateGuidance,
    '',
    'Quality expectations:',
    '- lessonSummary must be a substantial teaching overview, not a single sentence.',
    '- successCriteria must be student-friendly and measurable.',
    '- vocabulary terms must be age-appropriate.',
    '- teacherNotes must include teaching tips, pacing advice, or likely misconceptions.',
    '- Ensure the lesson is genuinely useful for classroom instruction, not generic filler.'
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
      '- Use direct explanation followed by scaffolded examples.',
      '- Practice items should start easier and gradually increase in complexity.'
    ].join('\n');
  }

  if (t === 'REVIEW') {
    return [
      '- Focus on concise recall, representative examples, and mixed review tasks.',
      '- Keep the summary focused and practical.'
    ].join('\n');
  }

  if (t === 'QUIZ') {
    return [
      '- Keep the lesson summary short and assessment-oriented.',
      '- Practice items should function like quiz items with clear answerability.'
    ].join('\n');
  }

  if (t === 'REMEDIATION') {
    return [
      '- Break concepts into smaller steps.',
      '- Use simpler wording and explicit scaffolds.'
    ].join('\n');
  }

  if (t === 'ENRICHMENT') {
    return [
      '- Include extension-level thinking, pattern recognition, or transfer when appropriate.',
      '- Practice items should include some higher-level application.'
    ].join('\n');
  }

  return [
    '- Build for direct instruction first, then guided or independent practice.',
    '- Ensure the lesson is classroom-ready and logically sequenced.'
  ].join('\n');
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

  const topHeaderHtml = renderTopHeaderHtml_(lesson, teacherView, headerInfo);
  const objectiveHtml = `
    <section>
      <h2 class="section-title">Objective</h2>
      <div class="card">
        <div class="editable objective-text" data-field="objective">${escapeHtml_(lesson.objective)}</div>
      </div>
    </section>
  `;

  const successCriteriaHtml = lesson.successCriteria && lesson.successCriteria.length ? `
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
  ` : '';

  const vocabularyHtml = lesson.vocabulary && lesson.vocabulary.length ? `
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
  ` : '';

  const summaryHtml = `
    <section>
      <h2 class="section-title">Lesson Overview</h2>
      <div class="card">
        <div class="section-note editable summary-text" data-field="lessonSummary">${escapeHtml_(lesson.lessonSummary)}</div>
      </div>
    </section>
  `;

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

  const practiceHtml = lesson.practiceItems.map(function (item, index) {
    return renderPracticeItemHtml_(item, index, teacherView, forPrint);
  }).join('');

  const teacherNotesHtml = teacherView ? `
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
  ` : '';

  const lessonSheetClass = 'lesson-sheet layout-' + layoutMode.toLowerCase();

  return `
    <div class="${lessonSheetClass}" data-copy-mode="${copyMode}">
      ${topHeaderHtml}
      ${objectiveHtml}
      ${successCriteriaHtml}
      ${vocabularyHtml}
      ${summaryHtml}

      <section>
        <h2 class="section-title">Key Concepts</h2>
        <div class="section-note">Use these explanations and examples for direct instruction.</div>
        ${conceptsHtml}
      </section>

      <section class="practice-section">
        <h2 class="section-title">Individual Practice</h2>
        <div class="section-note">
          ${teacherView
            ? 'Teacher view shows accepted answers and hints.'
            : 'Answer all items. Use the interactive buttons for immediate checking.'}
        </div>
        ${practiceHtml}
      </section>

      ${teacherNotesHtml}
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
    'body.print-mode .lesson-meta{grid-template-columns:repeat(auto-fit,minmax(1.7in,1fr));gap:8px;}',
    'body.print-mode .vocab-grid{grid-template-columns:repeat(auto-fit,minmax(2in,1fr));gap:8px;}',
    'body.print-mode .btn, body.print-mode .feedback, body.print-mode .answer-reveal.hidden{display:none !important;}',
    'body.print-mode .practice-controls{display:block;}',
    'body.print-mode .answer-input{display:block;width:100%;border:1px solid #999;min-width:0;}',
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


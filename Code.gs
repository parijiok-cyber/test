/**
 * 맞춤법 돋보기 (6모둠 · 넷이서 간다!)
 * Google Sheets + Apps Script 기반 웹앱
 *
 * - 학생 손글씨 사진을 업로드하면 Upstage AI가 글자를 읽고(Document Digitization OCR),
 *   Solar 챗 모델이 맞춤법/띄어쓰기 오류를 찾아 설명해준다.
 * - 오류가 없으면 학년에 맞는 추천 맞춤법(추가 학습 단어)을 보여준다.
 * - 학생들이 자주 틀리는 맞춤법을 스프레드시트에 누적하고, 그 데이터로 OX퀴즈/골든벨 퀴즈를 만든다.
 *
 * ============================ 배포 전 설정 (필독) ============================
 * 1. 이 프로젝트를 "Google Sheets에 바인딩된 Apps Script"로 만드는 것을 권장합니다.
 *    (스프레드시트 메뉴 확장 프로그램 > Apps Script)
 *    바인딩하지 않고 독립 스크립트로 쓰려면 SPREADSHEET_ID 스크립트 속성을 채워 넣으세요.
 * 2. 확장 프로그램 > Apps Script > 프로젝트 설정 > 스크립트 속성에 아래 키를 추가하세요.
 *      - UPSTAGE_API_KEY : console.upstage.ai 에서 발급받은 Upstage API 키 (up_로 시작)
 *      - (선택) UPSTAGE_CHAT_MODEL : 맞춤법 교정에 쓸 Solar 모델 ID. 비워두면 기본값 사용
 *      - (선택) SPREADSHEET_ID : 바인딩되지 않은 스크립트에서 사용할 스프레드시트 ID
 * 3. [배포] > [새 배포] > 유형: 웹 앱
 *      - 실행 계정: 나(교사)
 *      - 액세스 권한: 학교 조직 내 전체 또는 링크가 있는 모든 사용자 (학교 정책에 맞게 선택)
 *
 * 이 스크립트는 Upstage API를 두 번 호출합니다.
 *   ① Document Digitization(OCR) — 사진 속 손글씨를 텍스트로 읽어낸다.
 *   ② Chat Completions(Solar, function calling) — 읽어낸 텍스트에서 맞춤법 오류를 찾고 설명을 만든다.
 * ============================================================================
 */

// ---------------------------------------------------------------------------
// 설정
// ---------------------------------------------------------------------------

var SHEET_LOG = '오류기록';       // 학생별 맞춤법 오류 누적 로그
var SHEET_HEADER = [
  '타임스탬프', '학년', '학생(익명)', '원본 표현', '교정 표현', '오류 유형', '설명', '추출 원문'
];

var UPSTAGE_OCR_URL = 'https://api.upstage.ai/v1/document-digitization';
var UPSTAGE_OCR_MODEL = 'document-parse';

var UPSTAGE_CHAT_URL = 'https://api.upstage.ai/v1/chat/completions';
// 기본 Solar 모델. 스크립트 속성 UPSTAGE_CHAT_MODEL을 설정하면 그 값이 우선합니다.
// 최신 모델 ID는 Upstage 콘솔(console.upstage.ai)에서 확인해 바꿔주세요.
var UPSTAGE_CHAT_MODEL_DEFAULT = 'solar-pro2';

/**
 * 스크립트가 바인딩된 스프레드시트가 있으면 그것을, 없으면
 * 스크립트 속성 SPREADSHEET_ID로 지정된 스프레드시트를 연다.
 */
function getSpreadsheet_() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;

  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error(
      '스프레드시트를 찾을 수 없습니다. 이 스크립트를 스프레드시트에 바인딩하거나, ' +
      '스크립트 속성에 SPREADSHEET_ID를 설정해주세요.'
    );
  }
  return SpreadsheetApp.openById(id);
}

function getLogSheet_() {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.appendRow(SHEET_HEADER);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHEET_HEADER.length).setFontWeight('bold');
  }
  return sheet;
}

// ---------------------------------------------------------------------------
// 웹앱 라우팅
// ---------------------------------------------------------------------------

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) || 'main';
  var template;

  if (page === 'quiz') {
    template = HtmlService.createTemplateFromFile('QuizPage');
  } else {
    template = HtmlService.createTemplateFromFile('Index');
  }

  return template
    .evaluate()
    .setTitle('맞춤법 돋보기')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** HTML 파일 안에서 <?!= include('CSS'); ?> 형태로 다른 HTML 조각을 끼워 넣기 위한 헬퍼 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// 핵심 기능 1~3: 손글씨 인식 + 맞춤법 교정 + 추천 학습
// ---------------------------------------------------------------------------

/**
 * 클라이언트(Index.html)에서 호출.
 * @param {string} base64Data  이미지의 순수 base64 데이터 (data: 접두어 제외)
 * @param {string} mimeType    예: "image/jpeg"
 * @param {string} studentAlias 학생 익명 코드/번호 (예: "3-2-05" 등). 실명 금지.
 * @param {string} grade        학년 (예: "3학년"). 없으면 빈 문자열.
 * @return {Object} { extractedText, corrections:[{original,corrected,reason}], praise, recommendedWords:[{word,tip}] }
 */
function analyzeHandwriting(base64Data, mimeType, studentAlias, grade) {
  if (!base64Data) throw new Error('이미지 데이터가 없습니다.');

  var props = PropertiesService.getScriptProperties();
  var apiKey = props.getProperty('UPSTAGE_API_KEY');
  if (!apiKey) {
    throw new Error('UPSTAGE_API_KEY가 설정되어 있지 않습니다. 스크립트 속성을 확인해주세요.');
  }

  var gradeText = grade || '학년 정보 없음';

  // ① Document Digitization(OCR)으로 사진 속 손글씨를 텍스트로 읽는다.
  var extractedText = ocrExtractText_(base64Data, mimeType, apiKey);

  if (!extractedText || !extractedText.trim()) {
    throw new Error('사진에서 글씨를 읽어내지 못했어요. 더 밝고 또렷한 사진으로 다시 시도해주세요.');
  }

  // ② Solar 챗 모델(function calling)로 맞춤법 오류를 찾고 설명을 만든다.
  var correction = correctText_(extractedText, gradeText, apiKey);

  var parsed = {
    extractedText: extractedText,
    corrections: correction.corrections || [],
    recommendedWords: correction.recommendedWords || [],
    praise: correction.praise || '',
  };

  logCorrections_(parsed, studentAlias, grade);

  return parsed;
}

/** Upstage Document Digitization API로 이미지에서 텍스트를 추출한다. */
function ocrExtractText_(base64Data, mimeType, apiKey) {
  var ext = (mimeType.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, 'handwriting.' + ext);

  var result = fetchWithRetry_(UPSTAGE_OCR_URL, {
    method: 'post',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: {
      document: blob,
      model: UPSTAGE_OCR_MODEL,
      mode: 'standard',
      ocr: 'force', // 손글씨이므로 OCR을 항상 강제 실행
    },
    muteHttpExceptions: true,
  });

  if (result && result.error) {
    throw new Error('Upstage OCR 오류: ' + (result.error.message || JSON.stringify(result.error)));
  }

  var content = result && result.content;
  if (!content) return '';

  if (content.text) return content.text;
  if (content.markdown) {
    // 마크다운 기호(#, *, |, -- 등)를 걷어내 순수 텍스트에 가깝게 정리한다.
    return content.markdown
      .replace(/\|/g, ' ')
      .replace(/^[#>*\-\s]+/gm, '')
      .replace(/\n{2,}/g, '\n')
      .trim();
  }
  if (content.html) {
    return content.html.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim();
  }
  return '';
}

/** Upstage Solar 챗 모델(function calling)로 맞춤법 교정을 수행한다. */
function correctText_(extractedText, gradeText, apiKey) {
  var props = PropertiesService.getScriptProperties();
  var model = props.getProperty('UPSTAGE_CHAT_MODEL') || UPSTAGE_CHAT_MODEL_DEFAULT;

  var promptText =
    '당신은 친절하고 다정한 초등학교 국어 선생님입니다.\n' +
    '아래 텍스트는 학생(' + gradeText + ')이 손으로 쓴 글을 OCR로 읽어낸 것입니다.\n' +
    'OCR 과정에서 생긴 명백한 인식 오류(예: 없던 글자가 섞임)는 무시하고, ' +
    '실제 맞춤법·띄어쓰기·표준어 오류만 다루세요.\n\n' +
    '[학생이 쓴 글]\n' + extractedText + '\n\n' +
    '다음 작업을 수행한 뒤, 반드시 record_analysis 함수를 호출해서 결과를 전달하세요.\n' +
    '1. 텍스트에서 맞춤법, 띄어쓰기, 표준어 규정에 어긋난 부분을 찾아 교정한다.\n' +
    '2. 아이의 눈높이에 맞춰 왜 그렇게 고쳐야 하는지 친절하게 설명한다.\n' +
    '3. 만약 오류가 하나도 없다면, corrections는 빈 배열로 두고, 대신 ' + gradeText +
    ' 수준에서 자주 틀리는 맞춤법 단어를 recommendedWords에 2~3개 추천하여 추가 학습 자료로 제공한다.\n' +
    '4. 오류가 있다면 recommendedWords는 빈 배열로 둔다.';

  // OpenAI 호환 tools/tool_choice로 함수 호출을 강제해서, 안정적으로 구조화된 결과를 받는다.
  var tool = {
    type: 'function',
    function: {
      name: 'record_analysis',
      description: '학생 글쓰기의 맞춤법 분석 결과를 기록한다.',
      parameters: {
        type: 'object',
        properties: {
          corrections: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                original: { type: 'string', description: '틀리게 쓴 단어나 구절' },
                corrected: { type: 'string', description: '올바르게 고친 표현' },
                errorType: { type: 'string', description: '맞춤법 | 띄어쓰기 | 표준어' },
                reason: { type: 'string', description: '아이가 이해하기 쉬운 친절한 설명' },
              },
              required: ['original', 'corrected', 'reason'],
            },
          },
          recommendedWords: {
            type: 'array',
            items: {
              type: 'object',
              properties: {
                word: { type: 'string', description: '학년 수준에서 자주 틀리는 단어' },
                tip: { type: 'string', description: '쉬운 설명 한 줄' },
              },
            },
          },
          praise: { type: 'string', description: '글쓰기에 대한 따뜻한 칭찬 한마디' },
        },
        required: ['corrections', 'recommendedWords', 'praise'],
      },
    },
  };

  var payload = {
    model: model,
    messages: [{ role: 'user', content: promptText }],
    tools: [tool],
    tool_choice: { type: 'function', function: { name: 'record_analysis' } },
  };

  var result = fetchWithRetry_(UPSTAGE_CHAT_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  if (result && result.error) {
    throw new Error('Upstage Chat API 오류: ' + (result.error.message || JSON.stringify(result.error)));
  }

  var message = result && result.choices && result.choices[0] && result.choices[0].message;
  var toolCall = message && message.tool_calls && message.tool_calls[0];
  var argsText = toolCall && toolCall.function && toolCall.function.arguments;

  if (!argsText) {
    throw new Error('AI로부터 맞춤법 분석 결과를 받지 못했습니다. 잠시 후 다시 시도해주세요.');
  }

  return JSON.parse(argsText);
}

/** 재시도 로직이 포함된 UrlFetchApp 호출 (지수 백오프) */
function fetchWithRetry_(url, options, retries) {
  retries = retries || 4;
  var delays = [500, 1500, 3000, 6000];
  var lastError;

  for (var i = 0; i < retries; i++) {
    var response;
    try {
      response = UrlFetchApp.fetch(url, options);
    } catch (e) {
      lastError = e;
      if (i < retries - 1) Utilities.sleep(delays[i]);
      continue;
    }

    var code = response.getResponseCode();
    var body = response.getContentText();
    if (code >= 200 && code < 300) {
      return JSON.parse(body);
    }

    lastError = new Error('서버 응답 오류 (' + code + '): ' + body);

    // 429(과호출)나 5xx(서버 오류)만 재시도한다. 401/400 등은 키·요청 자체가
    // 잘못된 경우이므로 즉시 중단해서 학생을 불필요하게 오래 기다리게 하지 않는다.
    var isRetryable = code === 429 || code >= 500;
    if (!isRetryable) {
      throw lastError;
    }
    if (i < retries - 1) Utilities.sleep(delays[i]);
  }
  throw lastError;
}

/** 분석 결과를 '오류기록' 시트에 한 줄씩 누적 저장 */
function logCorrections_(parsed, studentAlias, grade) {
  var sheet = getLogSheet_();
  var now = new Date();
  var alias = studentAlias || '익명';
  var gradeText = grade || '';
  var extracted = parsed.extractedText || '';

  var corrections = parsed.corrections || [];
  if (corrections.length === 0) return;

  var rows = corrections.map(function (c) {
    return [
      now,
      gradeText,
      alias,
      c.original || '',
      c.corrected || '',
      c.errorType || '맞춤법',
      c.reason || '',
      extracted,
    ];
  });

  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, SHEET_HEADER.length)
    .setValues(rows);
}

// ---------------------------------------------------------------------------
// 핵심 기능 4: 자주 틀리는 맞춤법으로 OX퀴즈 / 골든벨 퀴즈 만들기
// ---------------------------------------------------------------------------

/**
 * '오류기록' 시트를 집계해서 (원본→교정) 쌍별 빈도수를 계산한다.
 * @param {number} limit 상위 몇 개를 가져올지
 */
function getFrequentErrors_(limit) {
  var sheet = getLogSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, SHEET_HEADER.length).getValues();
  var counts = {}; // key: original|||corrected -> {original, corrected, errorType, reason, count}

  data.forEach(function (row) {
    var original = row[3];
    var corrected = row[4];
    var errorType = row[5];
    var reason = row[6];
    if (!original || !corrected) return;

    var key = original + '|||' + corrected;
    if (!counts[key]) {
      counts[key] = { original: original, corrected: corrected, errorType: errorType, reason: reason, count: 0 };
    }
    counts[key].count += 1;
  });

  var list = Object.keys(counts).map(function (k) { return counts[k]; });
  list.sort(function (a, b) { return b.count - a.count; });

  return list.slice(0, limit || 10);
}

/**
 * 클라이언트(QuizPage.html)에서 호출.
 * OX퀴즈 형식: "다음 표현은 맞는 표현일까요?"
 *   - 절반은 실제 학생들이 자주 틀리는 "틀린 표현"을 그대로 보여주고 정답은 X
 *   - 나머지 절반은 그 "교정된(올바른) 표현"을 보여주고 정답은 O
 */
function generateOXQuiz(count) {
  var frequent = getFrequentErrors_(count || 8);
  if (frequent.length === 0) {
    return { questions: [], message: '아직 누적된 오류 데이터가 없습니다. 학생들의 글쓰기를 몇 건 더 분석해보세요.' };
  }

  var questions = frequent.map(function (item, idx) {
    var showWrong = idx % 2 === 0; // 번갈아가며 틀린 표현/맞는 표현 출제
    return {
      id: idx + 1,
      prompt: showWrong ? item.original : item.corrected,
      answer: showWrong ? 'X' : 'O',
      explanation: item.reason || (item.original + ' → ' + item.corrected),
      errorType: item.errorType || '맞춤법',
      count: item.count,
    };
  });

  return { questions: questions, message: '' };
}

/**
 * 골든벨 퀴즈: "다음 중 바르게 고친 표현은?" 4지선다
 */
function generateGoldenBellQuiz(count) {
  var frequent = getFrequentErrors_(count || 8);
  if (frequent.length === 0) {
    return { questions: [], message: '아직 누적된 오류 데이터가 없습니다. 학생들의 글쓰기를 몇 건 더 분석해보세요.' };
  }

  var allCorrected = frequent.map(function (item) { return item.corrected; });

  var questions = frequent.map(function (item, idx) {
    var distractors = shuffle_(
      allCorrected.filter(function (w) { return w !== item.corrected; })
    ).slice(0, 3);

    var options = shuffle_([item.corrected].concat(distractors));

    return {
      id: idx + 1,
      question: '"' + item.original + '"를(을) 바르게 고친 표현은 무엇일까요?',
      options: options,
      answer: item.corrected,
      explanation: item.reason || '',
      count: item.count,
    };
  });

  return { questions: questions, message: '' };
}

function shuffle_(array) {
  var a = array.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = a[i];
    a[i] = a[j];
    a[j] = tmp;
  }
  return a;
}

/** 대시보드용: 지금까지 누적된 오류 건수, 상위 오류 목록 */
function getDashboardSummary() {
  var sheet = getLogSheet_();
  var lastRow = sheet.getLastRow();
  var totalCount = Math.max(0, lastRow - 1);
  var top = getFrequentErrors_(5);
  return { totalCount: totalCount, top: top };
}

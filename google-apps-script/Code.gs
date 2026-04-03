/**
 * Google Apps Script — PRD 기획안 수집용 (1-page / 상세 / 피드백 통합)
 *
 * [설치 방법]
 * 1. 구글 시트 열기 → 확장 프로그램 → Apps Script
 * 2. 이 코드를 전체 복사하여 붙여넣기
 * 3. 상단 메뉴 "배포" → "새 배포" (또는 "배포 관리"에서 버전 업데이트)
 * 4. 유형: "웹 앱" 선택
 * 5. 실행 주체: "나" / 액세스 권한: "모든 사용자"
 * 6. 배포 후 생성된 URL을 복사
 * 7. HTML 파일의 SCRIPT_URL 변수에 붙여넣기
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // AI 확장 요청 처리
    if (data.action === 'ai_expand') {
      return handleAiExpand(data);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var action = data.action || 'submit';
    var type = data.type || '1page';

    // 탭 이름: "임시저장_1page", "제출완료_상세" 등
    var prefix = action === 'draft' ? '임시저장' : '제출완료';
    var typeName = getTypeName(type);
    var sheetName = prefix + '_' + typeName;

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      var headers = getHeaders(type);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#f0f0f0');
      sheet.setFrozenRows(1);
    }

    var existingRow = findExistingRow(sheet, data.meta_author, data.meta_project || data.meta_plan_name || '');
    var row = buildRow(data);

    if (existingRow > 0) {
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    for (var i = 1; i <= Math.min(row.length, 10); i++) {
      sheet.autoResizeColumn(i);
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        status: 'success',
        message: (action === 'draft' ? '임시저장' : '제출') + ' 완료',
        timestamp: new Date().toISOString()
      })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', message: 'PRD 수집 서버 작동 중' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// 타입 → 한글 이름
function getTypeName(type) {
  if (type === 'detailed') return '상세';
  if (type === 'feedback') return '피드백';
  return '1page';
}

// ═══════════════════════════════════════
// 헤더 정의
// ═══════════════════════════════════════
function getHeaders(type) {
  if (type === 'detailed') return getDetailedHeaders();
  if (type === 'feedback') return getFeedbackHeaders();
  return getOnepageHeaders();
}

function getOnepageHeaders() {
  return [
    '타임스탬프', '작성자', '부서', '프로젝트명',
    '제품 유형', '한 줄 정의', '문제 정의', '목표',
    'Primary 사용자', 'Primary 업무', 'Primary Pain Point',
    'Secondary 사용자', 'Secondary 업무', 'Secondary Pain Point',
    'F-01 기능명', 'F-01 설명', 'F-01 우선순위', 'F-01 AI활용',
    'F-02 기능명', 'F-02 설명', 'F-02 우선순위', 'F-02 AI활용',
    'F-03 기능명', 'F-03 설명', 'F-03 우선순위', 'F-03 AI활용',
    'F-04 기능명', 'F-04 설명', 'F-04 우선순위', 'F-04 AI활용',
    'AS-IS 프로세스', 'TO-BE 프로세스',
    'AS-IS 소요시간', 'TO-BE 소요시간',
    'AS-IS 품질', 'TO-BE 품질',
    'KPI 1', 'KPI 1 현재', 'KPI 1 목표', 'KPI 1 측정',
    'KPI 2', 'KPI 2 현재', 'KPI 2 목표', 'KPI 2 측정',
    'KPI 3', 'KPI 3 현재', 'KPI 3 목표', 'KPI 3 측정',
    'Phase 1 기간', 'Phase 1 산출물',
    'Phase 2 기간', 'Phase 2 산출물',
    'Phase 3 기간', 'Phase 3 산출물',
    '리스크 1', '대응 1',
    '리스크 2', '대응 2',
    '리스크 3', '대응 3',
    '활용 기술', '상태'
  ];
}

function getDetailedHeaders() {
  return [
    '타임스탬프', '작성자', '부서', '프로젝트명',
    'Executive Summary',
    '사업 환경', '문제 정의', 'SMART_S', 'SMART_M', 'SMART_A', 'SMART_R', 'SMART_T',
    'In-Scope', 'Out-of-Scope',
    'Primary 사용자', 'Primary 업무', 'Primary 니즈',
    'Secondary 사용자', 'Secondary 업무', 'Secondary 니즈',
    '페르소나1 이름', '페르소나1 업무', '페르소나1 Pain', '페르소나1 목표',
    '페르소나2 이름', '페르소나2 업무', '페르소나2 Pain', '페르소나2 목표',
    '시나리오1 Given', '시나리오1 When', '시나리오1 Then',
    '시나리오2 Given', '시나리오2 When', '시나리오2 Then',
    '시나리오3 Given', '시나리오3 When', '시나리오3 Then',
    'AS-IS Step1', 'AS-IS Step2', 'AS-IS Step3', 'AS-IS Step4',
    'TO-BE 소요시간 AS', 'TO-BE 소요시간 TO', 'TO-BE 소요시간 효과',
    'TO-BE 처리량 AS', 'TO-BE 처리량 TO', 'TO-BE 처리량 효과',
    'TO-BE 정확도 AS', 'TO-BE 정확도 TO', 'TO-BE 정확도 효과',
    'TO-BE 비용 AS', 'TO-BE 비용 TO', 'TO-BE 비용 효과',
    'F-01 기능명', 'F-01 설명', 'F-01 우선순위',
    'F-02 기능명', 'F-02 설명', 'F-02 우선순위',
    'F-03 기능명', 'F-03 설명', 'F-03 우선순위',
    'F-04 기능명', 'F-04 설명', 'F-04 우선순위',
    'F-05 기능명', 'F-05 설명', 'F-05 우선순위',
    'F-01 Input', 'F-01 Process', 'F-01 Output', 'F-01 룰',
    'F-02 Input', 'F-02 Process', 'F-02 Output', 'F-02 룰',
    '기술 프론트', '기술 백엔드', '기술 AI', '기술 DB', '기술 인프라',
    '외부 연동',
    'NFR 응답속도', 'NFR 동시접속', 'NFR 가용성', 'NFR 보안', 'NFR 규제', 'NFR 접근성',
    'KPI 1', 'KPI 1 현재', 'KPI 1 목표', 'KPI 1 측정',
    'KPI 2', 'KPI 2 현재', 'KPI 2 목표', 'KPI 2 측정',
    'KPI 3', 'KPI 3 현재', 'KPI 3 목표', 'KPI 3 측정',
    '정성 효과',
    '리스크 1', '리스크 1 영향', '리스크 1 확률', '리스크 1 대응',
    '리스크 2', '리스크 2 영향', '리스크 2 확률', '리스크 2 대응',
    '리스크 3', '리스크 3 영향', '리스크 3 확률', '리스크 3 대응',
    '상태'
  ];
}

function getFeedbackHeaders() {
  return [
    '타임스탬프', '검토자', '기획안명', '작성자',
    '점수_문제정의', '점수_AI적합성', '점수_사용자설계', '점수_기능요구',
    '점수_실행가능', '점수_성공지표', '점수_리스크', '점수_문서완성',
    '코멘트_문제정의', '코멘트_AI적합성', '코멘트_사용자설계', '코멘트_기능요구',
    '코멘트_실행가능', '코멘트_성공지표', '코멘트_리스크', '코멘트_문서완성',
    '총점', '등급',
    '강점 1', '강점 2', '강점 3',
    '개선점 1', '개선점 2', '개선점 3',
    '총평',
    '판정', '판정 사유',
    '액션 1', '액션 1 담당', '액션 1 기한',
    '액션 2', '액션 2 담당', '액션 2 기한',
    '액션 3', '액션 3 담당', '액션 3 기한',
    '상태'
  ];
}

// ═══════════════════════════════════════
// Row 빌드
// ═══════════════════════════════════════
function buildRow(data) {
  var type = data.type || '1page';
  if (type === 'detailed') return buildDetailedRow(data);
  if (type === 'feedback') return buildFeedbackRow(data);
  return buildOnepageRow(data);
}

function buildOnepageRow(data) {
  var f = data.fields || {};
  var row = [
    new Date(), data.meta_author || '', data.meta_dept || '', data.meta_project || '',
    f.product_type || '', f.oneliner || '', f.problem || '', f.objective || '',
    f.user_pri_type || '', f.user_pri_task || '', f.user_pri_pain || '',
    f.user_sec_type || '', f.user_sec_task || '', f.user_sec_pain || '',
  ];
  for (var i = 1; i <= 4; i++) {
    row.push(f['f' + i + '_name'] || '', f['f' + i + '_desc'] || '', f['f' + i + '_priority'] || '', f['f' + i + '_ai'] || '');
  }
  ['process', 'time', 'quality'].forEach(function(c) {
    row.push(f['asis_' + c] || '', f['tobe_' + c] || '');
  });
  for (var i = 1; i <= 3; i++) {
    row.push(f['kpi' + i + '_name'] || '', f['kpi' + i + '_current'] || '', f['kpi' + i + '_target'] || '', f['kpi' + i + '_method'] || '');
  }
  for (var i = 1; i <= 3; i++) {
    row.push(f['phase' + i + '_period'] || '', f['phase' + i + '_output'] || '');
  }
  for (var i = 1; i <= 3; i++) {
    row.push(f['risk' + i] || '', f['risk' + i + '_response'] || '');
  }
  row.push(f.tech_stack || '', data.action || 'submit');
  return row;
}

function buildDetailedRow(data) {
  var f = data.fields || {};
  var row = [new Date(), data.meta_author || '', data.meta_dept || '', data.meta_project || ''];

  // 순서대로 fields에서 꺼냄
  var keys = [
    'summary',
    'env', 'problem_detail', 'smart_s', 'smart_m', 'smart_a', 'smart_r', 'smart_t',
    'in_scope', 'out_scope',
    'user_pri_type', 'user_pri_task', 'user_pri_need',
    'user_sec_type', 'user_sec_task', 'user_sec_need',
    'p1_name', 'p1_task', 'p1_pain', 'p1_goal',
    'p2_name', 'p2_task', 'p2_pain', 'p2_goal',
    'sc1_given', 'sc1_when', 'sc1_then',
    'sc2_given', 'sc2_when', 'sc2_then',
    'sc3_given', 'sc3_when', 'sc3_then',
    'asis_step1', 'asis_step2', 'asis_step3', 'asis_step4',
    'cmp_time_as', 'cmp_time_to', 'cmp_time_eff',
    'cmp_vol_as', 'cmp_vol_to', 'cmp_vol_eff',
    'cmp_acc_as', 'cmp_acc_to', 'cmp_acc_eff',
    'cmp_cost_as', 'cmp_cost_to', 'cmp_cost_eff',
  ];
  keys.forEach(function(k) { row.push(f[k] || ''); });

  for (var i = 1; i <= 5; i++) {
    row.push(f['df' + i + '_name'] || '', f['df' + i + '_desc'] || '', f['df' + i + '_priority'] || '');
  }
  for (var i = 1; i <= 2; i++) {
    row.push(f['df' + i + '_input'] || '', f['df' + i + '_process'] || '', f['df' + i + '_output'] || '', f['df' + i + '_rule'] || '');
  }

  ['tech_front', 'tech_back', 'tech_ai', 'tech_db', 'tech_infra', 'external_api'].forEach(function(k) {
    row.push(f[k] || '');
  });
  ['nfr_speed', 'nfr_concurrent', 'nfr_uptime', 'nfr_security', 'nfr_regulation', 'nfr_access'].forEach(function(k) {
    row.push(f[k] || '');
  });
  for (var i = 1; i <= 3; i++) {
    row.push(f['kpi' + i + '_name'] || '', f['kpi' + i + '_current'] || '', f['kpi' + i + '_target'] || '', f['kpi' + i + '_method'] || '');
  }
  row.push(f.qualitative || '');
  for (var i = 1; i <= 3; i++) {
    row.push(f['risk' + i] || '', f['risk' + i + '_impact'] || '', f['risk' + i + '_prob'] || '', f['risk' + i + '_response'] || '');
  }
  row.push(data.action || 'submit');
  return row;
}

function buildFeedbackRow(data) {
  var f = data.fields || {};
  var row = [new Date(), data.meta_author || '', f.plan_name || '', f.plan_author || ''];

  // 점수 8개
  for (var i = 1; i <= 8; i++) row.push(f['score_' + i] || '');
  // 코멘트 8개
  for (var i = 1; i <= 8; i++) row.push(f['comment_' + i] || '');

  row.push(f.total_score || '', f.grade || '');

  for (var i = 1; i <= 3; i++) row.push(f['strength_' + i] || '');
  for (var i = 1; i <= 3; i++) row.push(f['improve_' + i] || '');

  row.push(f.overall || '');
  row.push(f.verdict || '', f.verdict_reason || '');

  for (var i = 1; i <= 3; i++) {
    row.push(f['action_' + i] || '', f['action_' + i + '_who'] || '', f['action_' + i + '_due'] || '');
  }

  row.push(data.action || 'submit');
  return row;
}

// ═══════════════════════════════════════
// AI 확장 (1page → 상세 기획안)
// ═══════════════════════════════════════
function handleAiExpand(data) {
  var apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: 'API 키가 설정되지 않았습니다. Script Properties에 OPENAI_API_KEY를 추가하세요.' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var f = data.fields || {};

  var userInput = '[프로젝트 정보]\n'
    + '프로젝트명: ' + (data.meta_project || '') + '\n'
    + '작성자: ' + (data.meta_author || '') + ' / 부서: ' + (data.meta_dept || '') + '\n\n'
    + '[제품 개요]\n'
    + '제품 유형: ' + (f.product_type || '') + '\n'
    + '한 줄 정의: ' + (f.oneliner || '') + '\n\n'
    + '[문제 정의 & 목표]\n'
    + '해결하려는 문제: ' + (f.problem || '') + '\n'
    + '목표: ' + (f.objective || '') + '\n\n'
    + '[대상 사용자]\n'
    + 'Primary: ' + (f.user_pri_type || '') + ' / ' + (f.user_pri_task || '') + ' / Pain: ' + (f.user_pri_pain || '') + '\n'
    + 'Secondary: ' + (f.user_sec_type || '') + ' / ' + (f.user_sec_task || '') + ' / Pain: ' + (f.user_sec_pain || '') + '\n\n'
    + '[핵심 기능]\n'
    + 'F-01: ' + (f.f1_name || '') + ' - ' + (f.f1_desc || '') + ' (' + (f.f1_priority || '') + ', AI: ' + (f.f1_ai || '') + ')\n'
    + 'F-02: ' + (f.f2_name || '') + ' - ' + (f.f2_desc || '') + ' (' + (f.f2_priority || '') + ', AI: ' + (f.f2_ai || '') + ')\n'
    + 'F-03: ' + (f.f3_name || '') + ' - ' + (f.f3_desc || '') + ' (' + (f.f3_priority || '') + ', AI: ' + (f.f3_ai || '') + ')\n'
    + 'F-04: ' + (f.f4_name || '') + ' - ' + (f.f4_desc || '') + ' (' + (f.f4_priority || '') + ', AI: ' + (f.f4_ai || '') + ')\n\n'
    + '[AS-IS / TO-BE]\n'
    + '프로세스: ' + (f.asis_process || '') + ' → ' + (f.tobe_process || '') + '\n'
    + '소요시간: ' + (f.asis_time || '') + ' → ' + (f.tobe_time || '') + '\n'
    + '품질: ' + (f.asis_quality || '') + ' → ' + (f.tobe_quality || '') + '\n\n'
    + '[KPI]\n'
    + 'KPI1: ' + (f.kpi1_name || '') + ' (현재: ' + (f.kpi1_current || '') + ', 목표: ' + (f.kpi1_target || '') + ')\n'
    + 'KPI2: ' + (f.kpi2_name || '') + ' (현재: ' + (f.kpi2_current || '') + ', 목표: ' + (f.kpi2_target || '') + ')\n'
    + 'KPI3: ' + (f.kpi3_name || '') + ' (현재: ' + (f.kpi3_current || '') + ', 목표: ' + (f.kpi3_target || '') + ')\n\n'
    + '[일정]\n'
    + 'Phase 1: ' + (f.phase1_period || '') + ' / ' + (f.phase1_output || '') + '\n'
    + 'Phase 2: ' + (f.phase2_period || '') + ' / ' + (f.phase2_output || '') + '\n'
    + 'Phase 3: ' + (f.phase3_period || '') + ' / ' + (f.phase3_output || '') + '\n\n'
    + '[리스크]\n'
    + '리스크1: ' + (f.risk1 || '') + ' → 대응: ' + (f.risk1_response || '') + '\n'
    + '리스크2: ' + (f.risk2 || '') + ' → 대응: ' + (f.risk2_response || '') + '\n'
    + '리스크3: ' + (f.risk3 || '') + ' → 대응: ' + (f.risk3_response || '') + '\n\n'
    + '[활용 기술]\n' + (f.tech_stack || '');

  var systemPrompt = '당신은 증권사의 AI 업무 효율화 PRD 상세 기획안 작성 도우미입니다.\n'
    + '1-Page Summary 데이터를 기반으로 상세 기획안에 필요한 추가 필드를 JSON으로 생성하세요.\n\n'
    + '[규칙]\n'
    + '1. 한국어로 작성\n'
    + '2. 금융/증권 업계 맥락에 맞게 구체적으로\n'
    + '3. 1page 내용을 확장/구체화 (기존 내용과 자연스럽게 연결)\n'
    + '4. 각 필드 1~3문장 (간결하되 구체적)\n'
    + '5. sketch_* 필드는 텍스트 ASCII 와이어프레임\n'
    + '6. JSON만 출력\n\n'
    + '[생성할 필드]\n'
    + 'summary: 프로젝트 3~5문장 Executive Summary\n'
    + 'env: 시장 상황, 경쟁 환경, 사업 배경 (2~3문장)\n'
    + 'problem_detail: 문제를 정량 근거와 함께 확장 (3~4문장)\n'
    + 'smart_s, smart_m, smart_a, smart_r, smart_t: SMART 목표 각 항목 (각 1문장)\n'
    + 'in_scope: 프로젝트 포함 범위 (항목 나열)\n'
    + 'out_scope: 제외 범위 (항목 나열)\n'
    + 'user_pri_need, user_sec_need: 사용자 핵심 니즈 (각 1문장)\n'
    + 'p1_name, p1_task, p1_pain, p1_goal: 페르소나1 (Primary 기반, "이름 / 역할 N년차" 형식)\n'
    + 'p2_name, p2_task, p2_pain, p2_goal: 페르소나2 (Secondary 기반)\n'
    + 'sc1_given, sc1_when, sc1_then: 시나리오1 Given-When-Then\n'
    + 'sc2_given, sc2_when, sc2_then: 시나리오2\n'
    + 'sc3_given, sc3_when, sc3_then: 시나리오3\n'
    + 'journey_act1~5: 유저 저니 행동 (인지→접근→사용→결과확인→재사용)\n'
    + 'journey_feel1~5: 유저 저니 감정\n'
    + 'journey_pain1~5: 유저 저니 Pain Point\n'
    + 'asis_step1~4: 현행 AS-IS 프로세스 4단계\n'
    + 'cmp_time_eff: 소요시간 개선 효과 (예: "87% 감소")\n'
    + 'cmp_vol_as, cmp_vol_to, cmp_vol_eff: 처리량 비교\n'
    + 'cmp_acc_eff: 정확도 개선 효과\n'
    + 'cmp_cost_as, cmp_cost_to, cmp_cost_eff: 비용 비교\n'
    + 'df5_name, df5_desc, df5_priority: 추가 5번째 기능 (기존과 중복 없이)\n'
    + 'df1_input, df1_process, df1_output, df1_rule: F-01 상세 I/P/O/규칙\n'
    + 'df2_input, df2_process, df2_output, df2_rule: F-02 상세\n'
    + 'sketch_main: 메인 화면 ASCII 와이어프레임\n'
    + 'sketch_detail: 핵심 기능 화면 와이어프레임\n'
    + 'sketch_architecture: 시스템 구성도 (데이터 흐름)\n'
    + 'tech_front, tech_back, tech_ai, tech_db, tech_infra: 기술 스택\n'
    + 'external_api: 외부 연동 시스템 (번호 목록)\n'
    + 'nfr_speed, nfr_concurrent, nfr_uptime, nfr_security, nfr_regulation, nfr_access: 비기능 요구사항\n'
    + 'phase1_milestone, phase2_milestone, phase3_milestone: 각 Phase 마일스톤\n'
    + 'qualitative: 정성적 효과 (2~3문장)\n'
    + 'risk1_impact, risk1_prob, risk2_impact, risk2_prob, risk3_impact, risk3_prob: H/M/L 값\n'
    + 'hist1_ver, hist1_date, hist1_desc, hist1_author: 문서 이력 초안 정보';

  var payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: '다음 1-Page Summary를 기반으로 상세 기획안 필드를 생성해주세요:\n\n' + userInput }
    ],
    temperature: 0.7,
    response_format: { type: 'json_object' }
  };

  var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    headers: {
      'Authorization': 'Bearer ' + apiKey,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: 'OpenAI API 오류 (' + code + ')' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var result = JSON.parse(response.getContentText());
  var content = JSON.parse(result.choices[0].message.content);

  return ContentService.createTextOutput(
    JSON.stringify({ status: 'success', fields: content })
  ).setMimeType(ContentService.MimeType.JSON);
}

// 기존 행 찾기
function findExistingRow(sheet, author, identifier) {
  if (!author) return -1;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === author && (identifier === '' || data[i][3] === identifier)) {
      return i + 1;
    }
  }
  return -1;
}

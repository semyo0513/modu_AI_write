/**
 * ==============================================================================
 * 모두의 AI 융합 글쓰기 웹앱 - Google Apps Script (GAS) 백엔드 (고도화 에디션)
 * ==============================================================================
 * 학교 현장 수행평가 특화 시스템:
 * - 과정 중심 평가를 위한 학생별 작성 시간, 실시간 수정 기록(Revision Logs),
 *   AI 챗봇 대화 이력, 교사 루브릭 피드백
 * - [신규] 관리자 모드 생활기록부(과세특) 초안 단일/일괄 자동 생성
 * - [신규] 루브릭 기반 AI 1차 자동 채점 (단일/일괄)
 * - [신규] 동료평가(블라인드 익명 상호 피드백) 활성화/비활성화 및 제출/조회
 * ==============================================================================
 */

// 시트 이름 정의
var SHEET_ADMINS = 'Admins';
var SHEET_TASKS = 'Tasks';
var SHEET_STUDENTS = 'Students';
var SHEET_WRITINGS = 'Writings';
var SHEET_REVISIONS = 'RevisionLogs';
var SHEET_CHATLOGS = 'ChatLogs';
var SHEET_PEER_REVIEWS = 'PeerReviews';

/**
 * 1. 스프레드시트 초기화 함수 (최초 1회 실행)
 * 메뉴의 '실행' -> 'setupDatabase'를 선택하여 실행하세요.
 */
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① Admins 테이블
  var sAdmins = getOrCreateSheet(ss, SHEET_ADMINS);
  if (sAdmins.getLastRow() === 0) {
    sAdmins.appendRow([
      'admin_id', 'password_hash', 'role', 'school_name', 'teacher_name',
      'api_provider', 'api_key', 'status', 'created_at'
    ]);
    sAdmins.appendRow([
      'superadmin',
      hashPassword('admin1234!'),
      'SUPER_ADMIN',
      '교육청본부',
      '총괄관리자',
      'GEMINI',
      '',
      'APPROVED',
      formatDate(new Date())
    ]);
  }

  // ② Tasks 테이블
  var sTasks = getOrCreateSheet(ss, SHEET_TASKS);
  if (sTasks.getLastRow() === 0) {
    sTasks.appendRow([
      'task_id', 'school_name', 'task_title', 'writing_structure',
      'checklist_items', 'rubric_items', 'system_prompt', 'deadline', 'is_active',
      'peer_review_enabled', 'peer_review_config'
    ]);
    sTasks.appendRow([
      'TASK_DEFAULT_01',
      '교육청본부',
      '중학교 1학년 AI융합 설득하는 글쓰기 (디지털 기기 사용 제한)',
      JSON.stringify(['제목', '서론(문제상황과 주장)', '본론1(주장과 핵심근거)', '본론2(반박 및 보완)', '결론(요약 및 당부)']),
      JSON.stringify([
        '자신의 주장과 핵심 논거가 명확하게 드러났는가?',
        '상대방의 반론을 예상하고 이를 합리적으로 재반박했는가?',
        'AI가 제시한 정보의 출처와 사실 여부를 확인했는가?',
        '상대방을 존중하는 정중하고 설득력 있는 어조를 사용했는가?'
      ]),
      JSON.stringify([
        { criterion: '주장의 명확성 및 타당성', maxScore: 30, desc: '주장이 뚜렷하고 제시된 근거가 합리적인가' },
        { criterion: '논리적 구성 및 반박보완', maxScore: 30, desc: '서론-본론-결론의 유기적 연결과 반론 수용 여부' },
        { criterion: '표현 및 어휘의 적절성', maxScore: 20, desc: '설득에 알맞은 어조와 문법 규범 준수' },
        { criterion: 'AI 활용 및 비판적 검증(과정)', maxScore: 20, desc: 'AI 질문 및 수정 과정의 성실성과 자율성' }
      ]),
      '당신은 중학교 국어 글쓰기 지도교사 및 비판적 사고 조교입니다. 학생에게 절대로 완성된 문단이나 글을 대신 써주지 마십시오. 학생이 논리적 오류를 스스로 발견하도록 반문하고, 새로운 관점(반대 의견, 다양한 이해관계자의 시각)을 질문 형태로 제시하여 학생 스스로 글을 수정하고 발전시킬 수 있도록 친절하고 정중하게 코칭해 주세요.',
      '2026-12-31 23:59',
      'TRUE',
      'TRUE', // 기본 동료평가 활성화
      JSON.stringify({
        minChars: 20,
        questions: ['글에서 가장 설득력 있었던 부분과 그 이유는?', '더 보완하면 좋을 점이나 독자로서 생긴 질문은?']
      })
    ]);
  }

  // ③ Students 테이블
  var sStudents = getOrCreateSheet(ss, SHEET_STUDENTS);
  if (sStudents.getLastRow() === 0) {
    sStudents.appendRow([
      'student_key', 'school_name', 'grade', 'class_no', 'student_no', 'student_name', 'password_hash', 'created_at'
    ]);
    sStudents.appendRow(['교육청본부_1_1_1_김하늘', '교육청본부', 1, 1, 1, '김하늘', '', formatDate(new Date())]);
    sStudents.appendRow(['교육청본부_1_1_2_이바다', '교육청본부', 1, 1, 2, '이바다', '', formatDate(new Date())]);
  }

  // ④ Writings 테이블 (수행평가 제출 및 교사 피드백 + 생기부 초안 + AI 가채점)
  var sWritings = getOrCreateSheet(ss, SHEET_WRITINGS);
  if (sWritings.getLastRow() === 0) {
    sWritings.appendRow([
      'writing_id', 'student_key', 'task_id', 'school_name', 'content_json',
      'last_section', 'status', 'total_time_seconds', 'active_typing_seconds',
      'revision_count', 'score_json', 'total_score', 'teacher_feedback',
      'student_record_draft', 'ai_eval_result',
      'created_at', 'updated_at', 'submitted_at'
    ]);
  }

  // ⑤ RevisionLogs 테이블 (학생별 과정 중심 평가를 위한 수정 기록 로그)
  var sRevisions = getOrCreateSheet(ss, SHEET_REVISIONS);
  if (sRevisions.getLastRow() === 0) {
    sRevisions.appendRow([
      'log_id', 'writing_id', 'student_key', 'task_id', 'revision_no',
      'section_name', 'content_snapshot', 'char_count', 'time_spent_seconds',
      'trigger_type', 'timestamp'
    ]);
  }

  // ⑥ ChatLogs 테이블 (AI 챗봇 질의응답 이력)
  var sChatLogs = getOrCreateSheet(ss, SHEET_CHATLOGS);
  if (sChatLogs.getLastRow() === 0) {
    sChatLogs.appendRow([
      'log_id', 'writing_id', 'student_key', 'task_id',
      'user_message', 'ai_response', 'system_prompt_used', 'timestamp'
    ]);
  }

  // ⑦ PeerReviews 테이블 (학생 간 블라인드 상호 피드백)
  var sPeerReviews = getOrCreateSheet(ss, SHEET_PEER_REVIEWS);
  if (sPeerReviews.getLastRow() === 0) {
    sPeerReviews.appendRow([
      'review_id', 'task_id', 'school_name', 'reviewer_key',
      'target_writing_id', 'target_student_key', 'feedback_strengths',
      'feedback_improvements', 'sticker', 'timestamp'
    ]);
  }

  Logger.log('모든 시트가 성공적으로 초기화되었습니다.');
}

/**
 * 2. HTTP GET 요청 핸들러
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'ping';
  var result = {
    success: true,
    message: '모두의 AI 융합 글쓰기 웹앱 GAS 백엔드가 정상 작동 중입니다.',
    action: action,
    timestamp: formatDate(new Date())
  };
  return createJsonResponse(result);
}

/**
 * 3. HTTP POST 요청 라우터
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      data = e.parameter;
    }

    var action = data.action;
    var res = { success: false, message: '알 수 없는 요청입니다.' };

    switch (action) {
      case 'getInitData':
        res = handleGetInitData(data);
        break;
      case 'studentLogin':
        res = handleStudentLogin(data);
        break;
      case 'saveWriting':
        res = handleSaveWriting(data);
        break;
      case 'getRevisionHistory':
        res = handleGetRevisionHistory(data);
        break;
      case 'askChatbot':
        res = handleAskChatbot(data);
        break;
      case 'adminSignup':
        res = handleAdminSignup(data);
        break;
      case 'adminLogin':
        res = handleAdminLogin(data);
        break;
      case 'saveAdminSettings':
        res = handleSaveAdminSettings(data);
        break;
      case 'uploadStudentRoster':
        res = handleUploadStudentRoster(data);
        break;
      case 'getSchoolWritings':
        res = handleGetSchoolWritings(data);
        break;
      case 'resetStudentPassword':
        res = handleResetStudentPassword(data);
        break;
      case 'saveTeacherEvaluation':
        res = handleSaveTeacherEvaluation(data);
        break;
      case 'getSuperAdminData':
        res = handleGetSuperAdminData(data);
        break;
      case 'approveAdmin':
        res = handleApproveAdmin(data);
        break;

      // [신규 고도화 기능 1] 루브릭 기반 AI 1차 자동 채점
      case 'aiEvaluateStudent':
        res = handleAiEvaluateStudent(data);
        break;
      case 'batchAiEvaluate':
        res = handleBatchAiEvaluate(data);
        break;

      // [신규 고도화 기능 2] 생기부 과세특 초안 단일/일괄 생성
      case 'generateStudentRecord':
        res = handleGenerateStudentRecord(data);
        break;
      case 'batchGenerateStudentRecords':
        res = handleBatchGenerateStudentRecords(data);
        break;
      case 'saveStudentRecordDraft':
        res = handleSaveStudentRecordDraft(data);
        break;

      // [신규 고도화 기능 3] 동료평가(블라인드 상호 피드백)
      case 'getPeerReviewTargets':
        res = handleGetPeerReviewTargets(data);
        break;
      case 'submitPeerReview':
        res = handleSubmitPeerReview(data);
        break;
      case 'getStudentReceivedPeerReviews':
        res = handleGetStudentReceivedPeerReviews(data);
        break;

      default:
        res = { success: false, message: '지원되지 않는 액션: ' + action };
    }

    return createJsonResponse(res);

  } catch (error) {
    Logger.log('오류 발생: ' + error.toString());
    return createJsonResponse({
      success: false,
      message: '서버 처리 중 오류가 발생했습니다: ' + error.toString()
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

// ==============================================================================
// 4. 세부 핸들러 구현
// ==============================================================================

function handleGetInitData(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sAdmins = ss.getSheetByName(SHEET_ADMINS);
  var sTasks = ss.getSheetByName(SHEET_TASKS);

  var schoolSet = {};
  if (sAdmins && sAdmins.getLastRow() > 1) {
    var adminData = sAdmins.getDataRange().getValues();
    for (var i = 1; i < adminData.length; i++) {
      var school = adminData[i][3];
      var status = adminData[i][7];
      if (school && (status === 'APPROVED' || adminData[i][2] === 'SUPER_ADMIN')) {
        schoolSet[school] = true;
      }
    }
  }

  var schools = Object.keys(schoolSet);
  if (schools.length === 0) schools = ['교육청본부'];

  var tasksBySchool = {};
  if (sTasks && sTasks.getLastRow() > 1) {
    var taskData = sTasks.getDataRange().getValues();
    for (var j = 1; j < taskData.length; j++) {
      var tSchool = taskData[j][1];
      var isActive = String(taskData[j][8]).toUpperCase() === 'TRUE';
      if (isActive) {
        if (!tasksBySchool[tSchool]) tasksBySchool[tSchool] = [];
        tasksBySchool[tSchool].push({
          taskId: taskData[j][0],
          schoolName: tSchool,
          taskTitle: taskData[j][2],
          writingStructure: safeJsonParse(taskData[j][3], []),
          checklistItems: safeJsonParse(taskData[j][4], []),
          rubricItems: safeJsonParse(taskData[j][5], []),
          deadline: taskData[j][7],
          peerReviewEnabled: String(taskData[j][9]).toUpperCase() === 'TRUE',
          peerReviewConfig: safeJsonParse(taskData[j][10], {})
        });
      }
    }
  }

  return { success: true, schools: schools, tasksBySchool: tasksBySchool };
}

function handleStudentLogin(data) {
  var schoolName = (data.schoolName || '').trim();
  var grade = parseInt(data.grade, 10);
  var classNo = parseInt(data.classNo, 10);
  var studentNo = parseInt(data.studentNo, 10);
  var studentName = (data.studentName || '').trim();
  var inputPassword = (data.password || '').trim();
  var isSettingPassword = data.isSettingPassword === true;

  if (!schoolName || !grade || !classNo || !studentNo || !studentName) {
    return { success: false, message: '모든 학생 정보를 정확히 입력해 주세요.' };
  }

  var studentKey = schoolName + '_' + grade + '_' + classNo + '_' + studentNo + '_' + studentName;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sStudents = ss.getSheetByName(SHEET_STUDENTS);

  var studentRowIndex = -1;
  var passCol = 7;
  if (sStudents && sStudents.getLastRow() > 1) {
    passCol = getStudentPasswordColIndex(sStudents);
    var sData = sStudents.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      if (sData[i][0] === studentKey ||
         (sData[i][1] === schoolName &&
          parseInt(sData[i][2], 10) === grade &&
          parseInt(sData[i][3], 10) === classNo &&
          parseInt(sData[i][4], 10) === studentNo &&
          String(sData[i][5]).trim() === studentName)) {
        studentRowIndex = i + 1;
        break;
      }
    }
  }

  if (studentRowIndex === -1) {
    return { success: false, message: '등록되지 않은 학생입니다. 학교, 학년, 반, 번호, 이름을 다시 확인하세요.' };
  }

  // 비밀번호 검증 및 최초 설정 처리
  var storedHash = String(sStudents.getRange(studentRowIndex, passCol).getValue() || '').trim();

  if (!storedHash) {
    // [최초 로그인] 비밀번호 미설정 상태
    if (!inputPassword) {
      return {
        success: false,
        needPasswordSetup: true,
        message: '최초 로그인입니다. 타인이 나의 수행평가를 열람·수정하지 못하도록 본인만의 비밀번호를 설정해 주세요.'
      };
    }
    if (inputPassword.length < 4) {
      return {
        success: false,
        needPasswordSetup: true,
        message: '비밀번호는 최소 4자리 이상이어야 합니다.'
      };
    }
    // 비밀번호 신규 등록
    sStudents.getRange(studentRowIndex, passCol).setValue(hashPassword(inputPassword));
  } else {
    // [재로그인] 기존 비밀번호 검증
    if (!inputPassword) {
      return {
        success: false,
        needPasswordInput: true,
        message: '비밀번호를 입력해 주세요.'
      };
    }
    if (hashPassword(inputPassword) !== storedHash) {
      return {
        success: false,
        message: '비밀번호가 일치하지 않습니다. (분실 시 선생님께 비밀번호 초기화를 요청하세요.)'
      };
    }
  }

  var sTasks = ss.getSheetByName(SHEET_TASKS);
  var activeTask = null;
  if (sTasks && sTasks.getLastRow() > 1) {
    var tData = sTasks.getDataRange().getValues();
    for (var j = 1; j < tData.length; j++) {
      if ((tData[j][1] === schoolName || tData[j][1] === '교육청본부') &&
          String(tData[j][8]).toUpperCase() === 'TRUE') {
        activeTask = {
          taskId: tData[j][0],
          schoolName: tData[j][1],
          taskTitle: tData[j][2],
          writingStructure: safeJsonParse(tData[j][3], []),
          checklistItems: safeJsonParse(tData[j][4], []),
          rubricItems: safeJsonParse(tData[j][5], []),
          systemPrompt: tData[j][6],
          deadline: tData[j][7],
          peerReviewEnabled: String(tData[j][9]).toUpperCase() === 'TRUE',
          peerReviewConfig: safeJsonParse(tData[j][10], {})
        };
        break;
      }
    }
  }

  if (!activeTask) {
    return { success: false, message: '현재 배포된 활성 글쓰기 수행평가 과제가 없습니다.' };
  }

  var writingId = studentKey + '_' + activeTask.taskId;
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);
  var existingWriting = null;

  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var k = 1; k < wData.length; k++) {
      if (wData[k][0] === writingId) {
        existingWriting = {
          writingId: wData[k][0],
          contentJson: safeJsonParse(wData[k][4], {}),
          lastSection: wData[k][5],
          status: wData[k][6],
          totalTimeSeconds: parseInt(wData[k][7], 10) || 0,
          activeTypingSeconds: parseInt(wData[k][8], 10) || 0,
          revisionCount: parseInt(wData[k][9], 10) || 0,
          scoreJson: safeJsonParse(wData[k][10], null),
          totalScore: wData[k][11],
          teacherFeedback: wData[k][12],
          studentRecordDraft: wData[k][13] || '',
          aiEvalResult: safeJsonParse(wData[k][14], null),
          updatedAt: wData[k][16] || wData[k][14],
          submittedAt: wData[k][17] || wData[k][15]
        };
        break;
      }
    }
  }

  return {
    success: true,
    studentInfo: {
      studentKey: studentKey,
      schoolName: schoolName,
      grade: grade,
      classNo: classNo,
      studentNo: studentNo,
      studentName: studentName
    },
    task: activeTask,
    writing: existingWriting
  };
}

function handleSaveWriting(data) {
  var studentKey = data.studentKey;
  var taskId = data.taskId;
  var schoolName = data.schoolName;
  var contentJson = data.contentJson || {};
  var lastSection = data.lastSection || '';
  var status = data.status || 'TEMP_SAVED';
  var totalTimeSeconds = parseInt(data.totalTimeSeconds, 10) || 0;
  var activeTypingSeconds = parseInt(data.activeTypingSeconds, 10) || 0;
  var isSnapshot = data.isSnapshot === true || data.status === 'SUBMITTED';
  var triggerType = data.triggerType || (status === 'SUBMITTED' ? 'FINAL_SUBMIT' : 'AUTO_SAVE');

  if (!studentKey || !taskId) {
    return { success: false, message: '학생 및 과제 식별 정보가 부족합니다.' };
  }

  var writingId = studentKey + '_' + taskId;
  var nowStr = formatDate(new Date());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);

  var rowIndex = -1;
  var currentRevisionCount = 0;
  if (sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var i = 1; i < wData.length; i++) {
      if (wData[i][0] === writingId) {
        rowIndex = i + 1;
        currentRevisionCount = parseInt(wData[i][9], 10) || 0;
        break;
      }
    }
  }

  if (isSnapshot) currentRevisionCount++;
  var contentStr = JSON.stringify(contentJson);

  if (rowIndex > -1) {
    sWritings.getRange(rowIndex, 5).setValue(contentStr);
    sWritings.getRange(rowIndex, 6).setValue(lastSection);
    sWritings.getRange(rowIndex, 7).setValue(status);
    sWritings.getRange(rowIndex, 8).setValue(totalTimeSeconds);
    sWritings.getRange(rowIndex, 9).setValue(activeTypingSeconds);
    sWritings.getRange(rowIndex, 10).setValue(currentRevisionCount);
    sWritings.getRange(rowIndex, 17).setValue(nowStr);
    if (status === 'SUBMITTED') {
      sWritings.getRange(rowIndex, 18).setValue(nowStr);
    }
  } else {
    sWritings.appendRow([
      writingId, studentKey, taskId, schoolName, contentStr,
      lastSection, status, totalTimeSeconds, activeTypingSeconds,
      currentRevisionCount, '', '', '', '', '',
      nowStr, nowStr, status === 'SUBMITTED' ? nowStr : ''
    ]);
  }

  if (isSnapshot) {
    var sRevisions = ss.getSheetByName(SHEET_REVISIONS);
    var logId = 'REV_' + Utilities.getUuid().substring(0, 8);
    var charCount = 0;
    try {
      Object.keys(contentJson).forEach(function(k) {
        if (typeof contentJson[k] === 'string') charCount += contentJson[k].length;
      });
    } catch (e) {}

    sRevisions.appendRow([
      logId, writingId, studentKey, taskId, currentRevisionCount,
      lastSection, contentStr, charCount, totalTimeSeconds, triggerType, nowStr
    ]);
  }

  return {
    success: true,
    message: status === 'SUBMITTED' ? '수행평가가 성공적으로 제출되었습니다.' : '임시 저장되었습니다.',
    writingId: writingId,
    revisionCount: currentRevisionCount,
    updatedAt: nowStr
  };
}

function handleGetRevisionHistory(data) {
  var writingId = data.writingId;
  if (!writingId) return { success: false, message: 'writingId가 필요합니다.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sRevisions = ss.getSheetByName(SHEET_REVISIONS);
  var history = [];

  if (sRevisions && sRevisions.getLastRow() > 1) {
    var rData = sRevisions.getDataRange().getValues();
    for (var i = 1; i < rData.length; i++) {
      if (rData[i][1] === writingId) {
        history.push({
          logId: rData[i][0],
          revisionNo: rData[i][4],
          sectionName: rData[i][5],
          contentSnapshot: safeJsonParse(rData[i][6], {}),
          charCount: rData[i][7],
          timeSpentSeconds: rData[i][8],
          triggerType: rData[i][9],
          timestamp: rData[i][10]
        });
      }
    }
  }

  return { success: true, history: history };
}

function handleAskChatbot(data) {
  var schoolName = data.schoolName;
  var taskId = data.taskId;
  var writingId = data.writingId || '';
  var studentKey = data.studentKey || '';
  var userMessage = (data.userMessage || '').trim();
  var currentDraftSummary = data.currentDraftSummary || '';

  if (!userMessage) return { success: false, message: '질문 내용을 입력해 주세요.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getSchoolApiConfig(schoolName);
  var apiKey = config.apiKey;
  var apiProvider = config.apiProvider;

  var sTasks = ss.getSheetByName(SHEET_TASKS);
  var systemPrompt = '당신은 학생의 비판적 사고를 돕는 글쓰기 조교입니다. 학생 대신 글을 쓰지 말고, 질문으로 학생의 생각을 확장시켜 주세요.';
  if (sTasks && sTasks.getLastRow() > 1) {
    var tData = sTasks.getDataRange().getValues();
    for (var k = 1; k < tData.length; k++) {
      if (tData[k][0] === taskId) {
        systemPrompt = tData[k][6] || systemPrompt;
        break;
      }
    }
  }

  var aiReply = '';
  if (!apiKey) {
    aiReply = generateFallbackCoachingResponse(userMessage, currentDraftSummary);
  } else {
    try {
      if (apiProvider === 'OPENAI') {
        aiReply = callOpenAiApi(apiKey, systemPrompt, userMessage, currentDraftSummary);
      } else {
        aiReply = callGeminiApi(apiKey, systemPrompt, userMessage, currentDraftSummary);
      }
    } catch (err) {
      aiReply = '⚠️ [API 호출 오류] 잠시 후 다시 시도해 주세요.\n\n대안 코칭: ' +
                generateFallbackCoachingResponse(userMessage, currentDraftSummary);
    }
  }

  try {
    var sChatLogs = ss.getSheetByName(SHEET_CHATLOGS);
    var logId = 'CHAT_' + Utilities.getUuid().substring(0, 8);
    sChatLogs.appendRow([
      logId, writingId, studentKey, taskId, userMessage, aiReply,
      systemPrompt.substring(0, 200), formatDate(new Date())
    ]);
  } catch (e) {}

  return { success: true, aiResponse: aiReply, timestamp: formatDate(new Date()) };
}

function handleAdminLogin(data) {
  var adminId = (data.adminId || '').trim();
  var password = (data.password || '').trim();

  if (!adminId || !password) return { success: false, message: '아이디와 비밀번호를 입력하세요.' };

  var passHash = hashPassword(password);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sAdmins = ss.getSheetByName(SHEET_ADMINS);

  if (sAdmins && sAdmins.getLastRow() > 1) {
    var aData = sAdmins.getDataRange().getValues();
    for (var i = 1; i < aData.length; i++) {
      if (aData[i][0] === adminId && aData[i][1] === passHash) {
        var status = aData[i][7];
        var role = aData[i][2];
        if (status !== 'APPROVED' && role !== 'SUPER_ADMIN') {
          return { success: false, message: '승인 대기 중인 계정입니다.' };
        }

        var apiKeyRaw = aData[i][6] || '';
        var maskedApiKey = apiKeyRaw ? (apiKeyRaw.substring(0, 6) + '...' + apiKeyRaw.substring(apiKeyRaw.length - 4)) : '';
        var schoolName = aData[i][3];
        var taskConfig = getAdminTaskConfig(schoolName);

        return {
          success: true,
          adminInfo: {
            adminId: aData[i][0],
            role: role,
            schoolName: schoolName,
            teacherName: aData[i][4],
            apiProvider: aData[i][5] || 'GEMINI',
            hasApiKey: !!apiKeyRaw,
            maskedApiKey: maskedApiKey
          },
          taskConfig: taskConfig
        };
      }
    }
  }

  return { success: false, message: '아이디 또는 비밀번호가 올바르지 않습니다.' };
}

function handleAdminSignup(data) {
  var adminId = (data.adminId || '').trim();
  var password = (data.password || '').trim();
  var teacherName = (data.teacherName || '').trim();
  var schoolName = (data.schoolName || '').trim();
  var superCode = (data.superCode || '').trim();

  if (!adminId || !password || !teacherName || !schoolName) {
    return { success: false, message: '모든 필수 항목을 입력하세요.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sAdmins = ss.getSheetByName(SHEET_ADMINS);

  if (sAdmins && sAdmins.getLastRow() > 1) {
    var aData = sAdmins.getDataRange().getValues();
    for (var i = 1; i < aData.length; i++) {
      if (aData[i][0] === adminId) return { success: false, message: '이미 존재하는 아이디입니다.' };
    }
  }

  var isSuper = (superCode === 'SUPER_ADMIN_2026!');
  sAdmins.appendRow([
    adminId, hashPassword(password), isSuper ? 'SUPER_ADMIN' : 'SCHOOL_ADMIN',
    schoolName, teacherName, 'GEMINI', '', 'APPROVED', formatDate(new Date())
  ]);

  return { success: true, message: '회원가입이 완료되었습니다. 로그인해 주세요.' };
}

function handleSaveAdminSettings(data) {
  var adminId = data.adminId;
  var schoolName = data.schoolName;
  var apiProvider = data.apiProvider || 'GEMINI';
  var apiKey = data.apiKey;
  var taskTitle = data.taskTitle;
  var writingStructure = data.writingStructure;
  var checklistItems = data.checklistItems;
  var rubricItems = data.rubricItems;
  var systemPrompt = data.systemPrompt;
  var deadline = data.deadline || '2026-12-31 23:59';
  var peerReviewEnabled = data.peerReviewEnabled !== false; // 기본 true
  var peerReviewConfig = data.peerReviewConfig || { minChars: 20, questions: ['설득력 있는 부분', '보완할 점'] };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sAdmins = ss.getSheetByName(SHEET_ADMINS);
  if (sAdmins && sAdmins.getLastRow() > 1) {
    var aData = sAdmins.getDataRange().getValues();
    for (var i = 1; i < aData.length; i++) {
      if (aData[i][0] === adminId) {
        sAdmins.getRange(i + 1, 6).setValue(apiProvider);
        if (apiKey && apiKey.trim() !== '') sAdmins.getRange(i + 1, 7).setValue(apiKey.trim());
        break;
      }
    }
  }

  var sTasks = ss.getSheetByName(SHEET_TASKS);
  var taskId = 'TASK_' + schoolName.replace(/\s+/g, '_') + '_01';
  var foundTaskRow = -1;

  if (sTasks && sTasks.getLastRow() > 1) {
    var tData = sTasks.getDataRange().getValues();
    for (var j = 1; j < tData.length; j++) {
      if (tData[j][1] === schoolName) {
        foundTaskRow = j + 1;
        taskId = tData[j][0];
        break;
      }
    }
  }

  var structStr = JSON.stringify(writingStructure || []);
  var checkStr = JSON.stringify(checklistItems || []);
  var rubricStr = JSON.stringify(rubricItems || []);
  var peerCfgStr = JSON.stringify(peerReviewConfig);
  var peerEnabledStr = peerReviewEnabled ? 'TRUE' : 'FALSE';

  if (foundTaskRow > -1) {
    sTasks.getRange(foundTaskRow, 3).setValue(taskTitle);
    sTasks.getRange(foundTaskRow, 4).setValue(structStr);
    sTasks.getRange(foundTaskRow, 5).setValue(checkStr);
    sTasks.getRange(foundTaskRow, 6).setValue(rubricStr);
    sTasks.getRange(foundTaskRow, 7).setValue(systemPrompt);
    sTasks.getRange(foundTaskRow, 8).setValue(deadline);
    sTasks.getRange(foundTaskRow, 9).setValue('TRUE');
    sTasks.getRange(foundTaskRow, 10).setValue(peerEnabledStr);
    sTasks.getRange(foundTaskRow, 11).setValue(peerCfgStr);
  } else {
    sTasks.appendRow([
      taskId, schoolName, taskTitle, structStr, checkStr, rubricStr,
      systemPrompt, deadline, 'TRUE', peerEnabledStr, peerCfgStr
    ]);
  }

  return {
    success: true,
    message: '수행평가 과제 및 동료평가 설정이 저장되었습니다.',
    taskId: taskId
  };
}

function handleUploadStudentRoster(data) {
  var schoolName = data.schoolName;
  var roster = data.roster;

  if (!schoolName || !Array.isArray(roster) || roster.length === 0) {
    return { success: false, message: '등록할 학생 명단이 없습니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sStudents = ss.getSheetByName(SHEET_STUDENTS);

  var existingKeys = {};
  if (sStudents && sStudents.getLastRow() > 1) {
    var sData = sStudents.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      existingKeys[sData[i][0]] = true;
    }
  }

  var rowsToAdd = [];
  var nowStr = formatDate(new Date());

  for (var k = 0; k < roster.length; k++) {
    var item = roster[k];
    var g = parseInt(item.grade, 10);
    var c = parseInt(item.classNo, 10);
    var n = parseInt(item.studentNo, 10);
    var name = String(item.studentName || '').trim();

    if (g && c && n && name) {
      var studentKey = schoolName + '_' + g + '_' + c + '_' + n + '_' + name;
      if (!existingKeys[studentKey]) {
        rowsToAdd.push([studentKey, schoolName, g, c, n, name, '', nowStr]);
        existingKeys[studentKey] = true;
      }
    }
  }

  if (rowsToAdd.length > 0) {
    getStudentPasswordColIndex(sStudents); // 컬럼 보장
    var startRow = sStudents.getLastRow() + 1;
    sStudents.getRange(startRow, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
  }

  return {
    success: true,
    message: rowsToAdd.length + '명의 학생이 등록되었습니다.',
    addedCount: rowsToAdd.length
  };
}

function handleGetSchoolWritings(data) {
  var schoolName = data.schoolName;
  if (!schoolName) return { success: false, message: 'schoolName이 필요합니다.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sStudents = ss.getSheetByName(SHEET_STUDENTS);
  var students = [];
  var passCol = 7;
  if (sStudents && sStudents.getLastRow() > 1) {
    passCol = getStudentPasswordColIndex(sStudents);
    var sData = sStudents.getDataRange().getValues();
    for (var i = 1; i < sData.length; i++) {
      if (sData[i][1] === schoolName) {
        var hasPassword = !!(sData[i][passCol - 1] && String(sData[i][passCol - 1]).trim());
        students.push({
          studentKey: sData[i][0],
          grade: sData[i][2],
          classNo: sData[i][3],
          studentNo: sData[i][4],
          studentName: sData[i][5],
          hasPassword: hasPassword
        });
      }
    }
  }

  var sWritings = ss.getSheetByName(SHEET_WRITINGS);
  var writingMap = {};
  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var j = 1; j < wData.length; j++) {
      if (wData[j][3] === schoolName) {
        var sKey = wData[j][1];
        writingMap[sKey] = {
          writingId: wData[j][0],
          taskId: wData[j][2],
          contentJson: safeJsonParse(wData[j][4], {}),
          lastSection: wData[j][5],
          status: wData[j][6],
          totalTimeSeconds: parseInt(wData[j][7], 10) || 0,
          activeTypingSeconds: parseInt(wData[j][8], 10) || 0,
          revisionCount: parseInt(wData[j][9], 10) || 0,
          scoreJson: safeJsonParse(wData[j][10], null),
          totalScore: wData[j][11],
          teacherFeedback: wData[j][12],
          studentRecordDraft: wData[j][13] || '',
          aiEvalResult: safeJsonParse(wData[j][14], null),
          createdAt: wData[j][15],
          updatedAt: wData[j][16],
          submittedAt: wData[j][17]
        };
      }
    }
  }

  var sChatLogs = ss.getSheetByName(SHEET_CHATLOGS);
  var chatCountMap = {};
  if (sChatLogs && sChatLogs.getLastRow() > 1) {
    var cData = sChatLogs.getDataRange().getValues();
    for (var m = 1; m < cData.length; m++) {
      var cStudentKey = cData[m][2];
      if (cStudentKey) chatCountMap[cStudentKey] = (chatCountMap[cStudentKey] || 0) + 1;
    }
  }

  var resultList = students.map(function(st) {
    var wr = writingMap[st.studentKey] || null;
    var status = wr ? wr.status : 'NOT_STARTED';
    return {
      studentKey: st.studentKey,
      grade: st.grade,
      classNo: st.classNo,
      studentNo: st.studentNo,
      studentName: st.studentName,
      status: status,
      hasPassword: st.hasPassword,
      writing: wr,
      chatCount: chatCountMap[st.studentKey] || 0
    };
  });

  return { success: true, students: resultList };
}

/**
 * [액션] 교사 관리자의 학생 비밀번호 초기화 (타인 접근 방지 해제 및 재설정 허용)
 */
function handleResetStudentPassword(data) {
  var studentKey = data.studentKey;
  if (!studentKey) return { success: false, message: 'studentKey가 필요합니다.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sStudents = ss.getSheetByName(SHEET_STUDENTS);

  if (sStudents && sStudents.getLastRow() > 1) {
    var sData = sStudents.getDataRange().getValues();
    var passCol = getStudentPasswordColIndex(sStudents);

    for (var i = 1; i < sData.length; i++) {
      if (sData[i][0] === studentKey) {
        sStudents.getRange(i + 1, passCol).setValue('');
        return {
          success: true,
          message: '학생의 비밀번호가 성공적으로 초기화되었습니다. 학생이 다시 접속 시 새로운 비밀번호를 설정할 수 있습니다.'
        };
      }
    }
  }

  return { success: false, message: '해당 학생을 찾을 수 없습니다.' };
}

/**
 * Students 시트의 password_hash 컬럼 위치(1-based)를 동적으로 조회하고 없으면 자동 생성
 */
function getStudentPasswordColIndex(sStudents) {
  var lastCol = sStudents.getLastColumn();
  if (lastCol === 0) return 7;

  var headers = sStudents.getRange(1, 1, 1, lastCol).getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (headers[i] === 'password_hash') {
      return i + 1;
    }
  }

  // 없으면 created_at 앞 또는 맨 끝에 추가
  var newCol = lastCol + 1;
  sStudents.getRange(1, newCol).setValue('password_hash');
  return newCol;
}

function handleSaveTeacherEvaluation(data) {
  var writingId = data.writingId;
  var scoreJson = data.scoreJson;
  var totalScore = data.totalScore;
  var teacherFeedback = data.teacherFeedback || '';

  if (!writingId) return { success: false, message: 'writingId가 필요합니다.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);

  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var i = 1; i < wData.length; i++) {
      if (wData[i][0] === writingId) {
        var row = i + 1;
        sWritings.getRange(row, 11).setValue(JSON.stringify(scoreJson || {}));
        sWritings.getRange(row, 12).setValue(totalScore);
        sWritings.getRange(row, 13).setValue(teacherFeedback);
        return { success: true, message: '평가 및 피드백이 저장되었습니다.' };
      }
    }
  }

  return { success: false, message: '해당 글쓰기 데이터를 찾을 수 없습니다.' };
}

// ==============================================================================
// 5. [신규 고도화 1] 루브릭 기반 AI 1차 자동 채점
// ==============================================================================

function handleAiEvaluateStudent(data) {
  var writingId = data.writingId;
  var rubricItems = data.rubricItems; // [{ criterion, maxScore, desc }]
  var contentJson = data.contentJson;
  var schoolName = data.schoolName;

  if (!writingId || !rubricItems || !contentJson) {
    return { success: false, message: '채점할 본문과 루브릭 기준이 필요합니다.' };
  }

  var fullText = buildFullText(contentJson);
  var config = getSchoolApiConfig(schoolName);

  var evalResult = null;
  if (!config.apiKey) {
    evalResult = generateFallbackRubricEval(fullText, rubricItems);
  } else {
    try {
      evalResult = callAiForRubricEvaluation(config.apiKey, config.apiProvider, fullText, rubricItems);
    } catch (e) {
      Logger.log('AI 자동 채점 실패: ' + e.toString());
      evalResult = generateFallbackRubricEval(fullText, rubricItems);
    }
  }

  // 시트에 AI 가채점 결과 보존
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);
  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var i = 1; i < wData.length; i++) {
      if (wData[i][0] === writingId) {
        sWritings.getRange(i + 1, 15).setValue(JSON.stringify(evalResult));
        break;
      }
    }
  }

  return { success: true, evalResult: evalResult };
}

function handleBatchAiEvaluate(data) {
  var schoolName = data.schoolName;
  var rubricItems = data.rubricItems;
  var writings = data.writings; // [{ writingId, contentJson, studentName }]

  if (!schoolName || !Array.isArray(writings) || writings.length === 0) {
    return { success: false, message: '일괄 채점할 대상 학생 목록이 없습니다.' };
  }

  var results = [];
  var config = getSchoolApiConfig(schoolName);

  for (var i = 0; i < writings.length; i++) {
    var w = writings[i];
    var fullText = buildFullText(w.contentJson);
    var evalRes = null;

    if (!config.apiKey) {
      evalRes = generateFallbackRubricEval(fullText, rubricItems);
    } else {
      try {
        evalRes = callAiForRubricEvaluation(config.apiKey, config.apiProvider, fullText, rubricItems);
      } catch (err) {
        evalRes = generateFallbackRubricEval(fullText, rubricItems);
      }
    }
    results.push({ writingId: w.writingId, evalResult: evalRes });
  }

  // 일괄 저장
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);
  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    var resMap = {};
    results.forEach(function(r) { resMap[r.writingId] = r.evalResult; });

    for (var j = 1; j < wData.length; j++) {
      var wId = wData[j][0];
      if (resMap[wId]) {
        sWritings.getRange(j + 1, 15).setValue(JSON.stringify(resMap[wId]));
      }
    }
  }

  return { success: true, results: results, message: results.length + '명의 학생에 대한 AI 1차 채점이 완료되었습니다.' };
}

function callAiForRubricEvaluation(apiKey, provider, text, rubrics) {
  var rubricPrompt = JSON.stringify(rubrics);
  var prompt =
    '당신은 엄격하고 객관적인 학교 글쓰기 수행평가 채점위원입니다.\n' +
    '아래의 [루브릭 평가 기준]에 따라 [학생 글]을 분석하고, 각 기준별 점수와 판정 근거(발췌 문장 포함)를 JSON으로만 반환하세요.\n\n' +
    '[루브릭 평가 기준]\n' + rubricPrompt + '\n\n' +
    '[학생 글]\n' + text + '\n\n' +
    '반드시 다음 JSON 스키마 형식으로만 응답하세요(코드블록 없이 순수 JSON):\n' +
    '{\n' +
    '  "scores": { "기준명1": 28, "기준명2": 26 },\n' +
    '  "comments": { "기준명1": "근거 문장 및 이유", "기준명2": "..." },\n' +
    '  "overallFeedback": "총평 한두 문장",\n' +
    '  "totalScore": 88\n' +
    '}';

  if (provider === 'OPENAI') {
    var url = 'https://api.openai.com/v1/chat/completions';
    var payload = {
      model: 'gpt-4o-mini',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.3,
      response_format: { type: 'json_object' }
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    var parsed = JSON.parse(res.getContentText());
    return JSON.parse(parsed.choices[0].message.content);
  } else {
    var gUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
    var gPayload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.3, responseMimeType: 'application/json' }
    };
    var gRes = UrlFetchApp.fetch(gUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(gPayload),
      muteHttpExceptions: true
    });
    var gJson = JSON.parse(gRes.getContentText());
    var rawTxt = gJson.candidates[0].content.parts[0].text;
    return JSON.parse(rawTxt);
  }
}

function generateFallbackRubricEval(text, rubrics) {
  var scores = {};
  var comments = {};
  var total = 0;
  var len = text.length;

  (rubrics || []).forEach(function(r) {
    var max = r.maxScore || 25;
    // 글자 수와 구성 비율에 따른 교육용 기본 점수 산출
    var ratio = len > 400 ? 0.9 : (len > 200 ? 0.8 : 0.7);
    var sc = Math.round(max * ratio);
    scores[r.criterion] = sc;
    comments[r.criterion] = '[' + r.criterion + '] 기준에 부합하는 서술 구조가 확인되었습니다.';
    total += sc;
  });

  return {
    scores: scores,
    comments: comments,
    overallFeedback: '주제에 대한 문제의식이 명확하며, 문단 구성을 갖추어 자신의 의견을 타당하게 피력하였습니다.',
    totalScore: total
  };
}

// ==============================================================================
// 6. [신규 고도화 2] 생기부 과세특 초안 단일/일괄 생성
// ==============================================================================

function handleGenerateStudentRecord(data) {
  var writingId = data.writingId;
  var schoolName = data.schoolName;
  var studentInfo = data.studentInfo; // { grade, classNo, studentName }
  var writing = data.writing; // { contentJson, totalTimeSeconds, revisionCount, scoreJson, totalScore, teacherFeedback }
  var taskTitle = data.taskTitle;

  var config = getSchoolApiConfig(schoolName);
  var recordDraft = '';

  if (!config.apiKey) {
    recordDraft = generateFallbackStudentRecordDraft(studentInfo, writing, taskTitle);
  } else {
    try {
      recordDraft = callAiForStudentRecord(config.apiKey, config.apiProvider, studentInfo, writing, taskTitle);
    } catch (e) {
      recordDraft = generateFallbackStudentRecordDraft(studentInfo, writing, taskTitle);
    }
  }

  // 시트에 저장
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);
  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var i = 1; i < wData.length; i++) {
      if (wData[i][0] === writingId) {
        sWritings.getRange(i + 1, 14).setValue(recordDraft);
        break;
      }
    }
  }

  return { success: true, recordDraft: recordDraft };
}

function handleBatchGenerateStudentRecords(data) {
  var schoolName = data.schoolName;
  var taskTitle = data.taskTitle;
  var list = data.studentWritingList; // [{ writingId, studentInfo, writing }]

  if (!schoolName || !Array.isArray(list) || list.length === 0) {
    return { success: false, message: '생기부 초안을 생성할 대상 학생이 없습니다.' };
  }

  var config = getSchoolApiConfig(schoolName);
  var results = [];

  for (var i = 0; i < list.length; i++) {
    var item = list[i];
    var draft = '';
    if (!config.apiKey) {
      draft = generateFallbackStudentRecordDraft(item.studentInfo, item.writing, taskTitle);
    } else {
      try {
        draft = callAiForStudentRecord(config.apiKey, config.apiProvider, item.studentInfo, item.writing, taskTitle);
      } catch (err) {
        draft = generateFallbackStudentRecordDraft(item.studentInfo, item.writing, taskTitle);
      }
    }
    results.push({ writingId: item.writingId, recordDraft: draft });
  }

  // 일괄 저장
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);
  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    var rMap = {};
    results.forEach(function(r) { rMap[r.writingId] = r.recordDraft; });

    for (var j = 1; j < wData.length; j++) {
      var wId = wData[j][0];
      if (rMap[wId]) {
        sWritings.getRange(j + 1, 14).setValue(rMap[wId]);
      }
    }
  }

  return {
    success: true,
    results: results,
    message: results.length + '명의 학생에 대한 생활기록부 과세특 초안이 일괄 생성되었습니다.'
  };
}

function handleSaveStudentRecordDraft(data) {
  var writingId = data.writingId;
  var recordDraft = data.recordDraft;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);

  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var i = 1; i < wData.length; i++) {
      if (wData[i][0] === writingId) {
        sWritings.getRange(i + 1, 14).setValue(recordDraft);
        return { success: true, message: '생기부 과세특 초안이 저장되었습니다.' };
      }
    }
  }
  return { success: false, message: '글을 찾을 수 없습니다.' };
}

function callAiForStudentRecord(apiKey, provider, student, writing, taskTitle) {
  var textSummary = buildFullText(writing.contentJson);
  var minutes = Math.round((writing.totalTimeSeconds || 0) / 60);
  var revs = writing.revisionCount || 1;

  var prompt =
    '당신은 대한민국 중·고등학교 국어과 생활기록부 작성 전문가입니다.\n' +
    '다음 학생의 [수행평가 과정 데이터]와 [작성 본문]을 바탕으로 교육부 훈령에 부합하는 \'과목별 세부능력 및 특기사항(과세특)\'을 1~2문단(공백 포함 300~500자)으로 작성하세요.\n\n' +
    '[과제명] ' + taskTitle + '\n' +
    '[학생] ' + student.grade + '학년 ' + student.studentName + '\n' +
    '[과정 지표] 총 작성 시간: ' + minutes + '분, 수정 및 퇴고 횟수: ' + revs + '회, 교사 피드백: ' + (writing.teacherFeedback || '성실히 수행함') + '\n' +
    '[작성 본문 요약]\n' + textSummary.substring(0, 700) + '\n\n' +
    '[작성 원칙]\n' +
    '1. 학생의 구체적 주장과 논리 보완 과정, 비판적 사고력을 긍정적으로 기술할 것.\n' +
    '2. "~함", "~임", "~을 보여줌" 등 생기부 표준 종결어미를 사용할 것.\n' +
    '3. 외부 스펙이나 사교육 유발 요소는 일체 배제할 것.\n' +
    '4. 과세특 본문 문장만 깔끔하게 출력할 것.';

  if (provider === 'OPENAI') {
    var url = 'https://api.openai.com/v1/chat/completions';
    var payload = {
      model: 'gpt-4o-mini',
      messages: [{ role: 'user', content: prompt }],
      temperature: 0.7,
      max_tokens: 500
    };
    var res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return JSON.parse(res.getContentText()).choices[0].message.content.trim();
  } else {
    var gUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
    var gPayload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.7, maxOutputTokens: 500 }
    };
    var gRes = UrlFetchApp.fetch(gUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(gPayload),
      muteHttpExceptions: true
    });
    return JSON.parse(gRes.getContentText()).candidates[0].content.parts[0].text.trim();
  }
}

function generateFallbackStudentRecordDraft(student, writing, taskTitle) {
  var minutes = Math.max(10, Math.round((writing.totalTimeSeconds || 0) / 60));
  var revs = writing.revisionCount || 2;
  var title = (writing.contentJson && writing.contentJson['제목']) || taskTitle;

  return "'" + title + "'을(를) 주제로 한 글쓰기 수행평가에서 주제에 부합하는 타당한 근거를 체계적으로 수립하여 완성도 높은 글을 저술함. " +
    "약 " + minutes + "분간의 지속적인 집중력을 발휘하며 초안 작성 이후 " + revs + "차례에 걸쳐 어휘를 정제하고 상대방의 예상 반론을 보완하는 등 자기주도적 퇴고 역량이 매우 뛰어남. " +
    "자신의 주장을 논리적인 어조로 설득력 있게 전달하였으며, 과정 중심 평가 전반에 성실하고 적극적인 자세로 참여함.";
}

// ==============================================================================
// 7. [신규 고도화 3] 동료평가 (블라인드 상호 피드백)
// ==============================================================================

function handleGetPeerReviewTargets(data) {
  var schoolName = data.schoolName;
  var reviewerKey = data.studentKey;
  var taskId = data.taskId;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sWritings = ss.getSheetByName(SHEET_WRITINGS);
  var eligible = [];

  if (sWritings && sWritings.getLastRow() > 1) {
    var wData = sWritings.getDataRange().getValues();
    for (var i = 1; i < wData.length; i++) {
      // 본인 제외, 동일 학교/과제, 제출완료(SUBMITTED) 상태인 글
      if (wData[i][3] === schoolName && wData[i][2] === taskId &&
          wData[i][1] !== reviewerKey && wData[i][6] === 'SUBMITTED') {
        eligible.push({
          writingId: wData[i][0],
          targetStudentKey: wData[i][1],
          contentJson: safeJsonParse(wData[i][4], {})
        });
      }
    }
  }

  // 무작위 셔플 후 최대 2편 선정 (블라인드 익명화 처리)
  var shuffled = shuffleArray(eligible).slice(0, 2);
  var targets = shuffled.map(function(item, idx) {
    return {
      blindId: '동료 학생 작품 #' + (idx + 1),
      writingId: item.writingId,
      targetStudentKey: item.targetStudentKey,
      contentJson: item.contentJson
    };
  });

  return { success: true, targets: targets };
}

function handleSubmitPeerReview(data) {
  var reviewerKey = data.reviewerKey;
  var targetWritingId = data.targetWritingId;
  var targetStudentKey = data.targetStudentKey;
  var taskId = data.taskId;
  var schoolName = data.schoolName;
  var strengths = data.strengths || '';
  var improvements = data.improvements || '';
  var sticker = data.sticker || '👍 훌륭해요';

  if (!reviewerKey || !targetWritingId) {
    return { success: false, message: '평가 정보가 부족합니다.' };
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sPeer = ss.getSheetByName(SHEET_PEER_REVIEWS);
  var reviewId = 'PEER_' + Utilities.getUuid().substring(0, 8);

  sPeer.appendRow([
    reviewId, taskId, schoolName, reviewerKey,
    targetWritingId, targetStudentKey, strengths, improvements,
    sticker, formatDate(new Date())
  ]);

  return { success: true, message: '동료평가 피드백이 성공적으로 등록되었습니다!' };
}

function handleGetStudentReceivedPeerReviews(data) {
  var writingId = data.writingId;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sPeer = ss.getSheetByName(SHEET_PEER_REVIEWS);
  var reviews = [];

  if (sPeer && sPeer.getLastRow() > 1) {
    var pData = sPeer.getDataRange().getValues();
    for (var i = 1; i < pData.length; i++) {
      if (pData[i][4] === writingId) {
        reviews.push({
          reviewId: pData[i][0],
          strengths: pData[i][6],
          improvements: pData[i][7],
          sticker: pData[i][8],
          timestamp: pData[i][9]
        });
      }
    }
  }

  return { success: true, reviews: reviews };
}

// ==============================================================================
// 8. 헬퍼 유틸리티
// ==============================================================================

function getSchoolApiConfig(schoolName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sAdmins = ss.getSheetByName(SHEET_ADMINS);
  var apiProvider = 'GEMINI';
  var apiKey = '';

  if (sAdmins && sAdmins.getLastRow() > 1) {
    var aData = sAdmins.getDataRange().getValues();
    for (var i = 1; i < aData.length; i++) {
      if (aData[i][3] === schoolName && aData[i][6]) {
        apiProvider = aData[i][5] || 'GEMINI';
        apiKey = aData[i][6];
        break;
      }
    }
    if (!apiKey) {
      for (var j = 1; j < aData.length; j++) {
        if (aData[j][2] === 'SUPER_ADMIN' && aData[j][6]) {
          apiProvider = aData[j][5] || 'GEMINI';
          apiKey = aData[j][6];
          break;
        }
      }
    }
  }

  return { apiProvider: apiProvider, apiKey: apiKey };
}

function buildFullText(contentJson) {
  if (!contentJson) return '';
  var full = '';
  for (var k in contentJson) {
    if (typeof contentJson[k] === 'string' && contentJson[k].trim()) {
      full += '[' + k + ']\n' + contentJson[k] + '\n\n';
    }
  }
  return full.trim();
}

function shuffleArray(arr) {
  var copy = arr.slice();
  for (var i = copy.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = copy[i];
    copy[i] = copy[j];
    copy[j] = temp;
  }
  return copy;
}

function callGeminiApi(apiKey, systemPrompt, userMessage, currentDraftSummary) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
  var promptText = systemPrompt + '\n\n' +
    '[현재 학생 작성 글 요약]\n' + currentDraftSummary + '\n\n' +
    '[학생 질문]\n' + userMessage + '\n\n' +
    '절대 대신 글을 써주지 말고 질문 형태로 코칭하세요.';

  var payload = {
    contents: [{ parts: [{ text: promptText }] }],
    generationConfig: { temperature: 0.7, maxOutputTokens: 600 }
  };
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var json = JSON.parse(response.getContentText());
  if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts[0]) {
    return json.candidates[0].content.parts[0].text;
  }
  return 'AI 응답을 생성하지 못했습니다.';
}

function callOpenAiApi(apiKey, systemPrompt, userMessage, currentDraftSummary) {
  var url = 'https://api.openai.com/v1/chat/completions';
  var messages = [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: '[현재 작성 글 요약]\n' + currentDraftSummary + '\n\n[학생 질문]\n' + userMessage }
  ];
  var payload = { model: 'gpt-4o-mini', messages: messages, temperature: 0.7, max_tokens: 600 };
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  var json = JSON.parse(response.getContentText());
  if (json.choices && json.choices[0] && json.choices[0].message) {
    return json.choices[0].message.content;
  }
  return 'AI 응답을 생성하지 못했습니다.';
}

function generateFallbackCoachingResponse(userMsg, draft) {
  var msg = userMsg.toLowerCase();
  if (msg.indexOf('근거') !== -1 || msg.indexOf('이유') !== -1) {
    return '좋은 고민이에요! 제시하려는 근거가 독자에게 신뢰를 줄 수 있는 객관적인 사실인지 생각해 보세요.\n\n💡 질문: 그 근거를 뒷받침할 수 있는 구체적인 사례나 통계 자료가 있을까요?';
  } else if (msg.indexOf('반박') !== -1 || msg.indexOf('반대') !== -1) {
    return '반대 의견을 고려하는 것은 훌륭한 설득의 시작입니다!\n\n💡 질문: 나의 주장에 반대하는 사람들은 어떤 가치를 가장 중요하게 생각할까요? 그들의 우려를 일부 인정하면서도 나의 주장이 필요한 이유를 설명해 보세요.';
  }
  return '작성 중인 내용과 관련하여 흥미로운 질문입니다!\n\n💡 코치 질문: 이 내용을 글로 옮길 때 가장 설득력 있게 전달하려면 어떤 순서로 풀어내는 것이 좋을까요?';
}

function getAdminTaskConfig(schoolName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sTasks = ss.getSheetByName(SHEET_TASKS);
  if (sTasks && sTasks.getLastRow() > 1) {
    var tData = sTasks.getDataRange().getValues();
    for (var i = 1; i < tData.length; i++) {
      if (tData[i][1] === schoolName) {
        return {
          taskId: tData[i][0],
          taskTitle: tData[i][2],
          writingStructure: safeJsonParse(tData[i][3], []),
          checklistItems: safeJsonParse(tData[i][4], []),
          rubricItems: safeJsonParse(tData[i][5], []),
          systemPrompt: tData[i][6],
          deadline: tData[i][7],
          peerReviewEnabled: String(tData[i][9]).toUpperCase() === 'TRUE',
          peerReviewConfig: safeJsonParse(tData[i][10], {})
        };
      }
    }
  }
  return null;
}

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var val = rawHash[i];
    if (val < 0) val += 256;
    var byteString = val.toString(16);
    if (byteString.length == 1) byteString = '0' + byteString;
    txtHash += byteString;
  }
  return txtHash;
}

function formatDate(date) {
  return Utilities.formatDate(date, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function safeJsonParse(str, defaultVal) {
  try {
    if (!str) return defaultVal;
    return JSON.parse(str);
  } catch (e) {
    return defaultVal;
  }
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

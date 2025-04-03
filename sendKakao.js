// sendKakao.gs

// 카톡으로 알림 보내기
function sendKakao() {

  try {
    // 1️⃣ 엑세스 토큰 갱신
    Logger.log("🔵 Access Token 갱신 시작");
    var newAccessToken = getNewAccessToken();  // getNewAccessToken()을 호출하여 새 토큰을 받는다.
    
    if (!newAccessToken) {
      Logger.log("🔴 Access Token 갱신 실패. 함수 종료.");
      return;
    }

    Logger.log("🟢 Access Token 갱신 완료: " + newAccessToken);
  
    // 2️⃣ 시트에서 데이터 가져오기
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PM 스케쥴");
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues(); // 시트값 가져오기
    var notes = dataRange.getNotes(); // 메모값 가져오기
    var mergedRanges = dataRange.getMergedRanges(); // 병합된 범위 가져오기

    // 3️⃣ 병합된 셀 데이터를 처리하여 각 행에 연결
    var processedData = processMergedRanges(data, notes, mergedRanges);

    // 4️⃣ 팀별 데이터 처리
    const teamCount = 2; // 총 팀 개수
    for (let i = 0; i < teamCount; i++) {
      processTeamData(processedData, `담당 ${i+1}팀`, i*10);
    }

    Logger.log("🎉 sendKakao() 실행 완료");

  } catch (e) {
    Logger.log("🔴 sendKakao 실행 중 오류 발생: " + e.message);
  }
}

// 병합된 셀 데이터 처리 함수
function processMergedRanges(data, notes, mergedRanges) {
  var processedData = JSON.parse(JSON.stringify(data)); // 원본 데이터를 복사하여 사용

  // 병합된 범위 처리
  mergedRanges.forEach(function(range) {
    var startRow = range.getRow();
    var endRow = startRow + range.getNumRows() - 1; // 병합 범위의 마지막 행
    var startColumn = range.getColumn();
    var value = range.getValue(); // 병합된 셀의 값
    var note = notes[startRow - 1][startColumn - 1] || "";

    // 병합 범위 내 모든 행에 동일한 값을 적용
    for (var row = startRow; row <= endRow; row++) {
      if (!processedData[row - 1]) continue; // 데이터가 없는 행은 건너뜀

      // 병합된 값을 해당 행에 삽입
      processedData[row - 1][startColumn - 1] = value;
      // 병합된 값(메모)을 해당 행에 삽입
      processedData[row - 1][startColumn] = note;
    }
  });

  return processedData;
}

// 팀별 데이터 처리 함수 (날짜만 비교)
function processTeamData(data, teamName, startColumn) {
  var today = new Date();
  today.setHours(0, 0, 0, 0); // 오늘 날짜의 시간을 00:00:00으로 설정

  for (var i = 1; i < data.length; i++) {
    var manager = data[i][startColumn]; // 담당자 이름
    var managerUUID = data[i][startColumn + 1]; // 담당자 토큰 값
    if (!manager) continue; // 담당자가 없으면 건너뜀

    var project = data[i][startColumn + 2]; // 프로젝트 이름
    var deadlines = {
      "기획": new Date(data[i][startColumn + 3]),
      "디자인": new Date(data[i][startColumn + 4]),
      "UI": new Date(data[i][startColumn + 5]),
      "기능": new Date(data[i][startColumn + 6]),
      "배포": new Date(data[i][startColumn + 7])
    };

    // Logger.log(`담당자: ${manager}, 프로젝트: ${project}, UUID: ${managerUUID}`);

    // 날짜 하루 전인지 확인 후 메시지 전송
    for (var task in deadlines) {
      var dueDate = deadlines[task];
      if (isNaN(dueDate)) continue; // 날짜가 유효하지 않으면 건너뜀
      dueDate.setHours(0, 0, 0, 0); // 마감일도 00:00:00으로 설정

      if ((dueDate - today) / (1000 * 60 * 60 * 24) === 1) {
        // 팀장은 전체 일정 발송 (팀명, 매니저명, UUID, 프로젝트명, task, dueDate)
        if (teamName == "담당 1팀" && manager == "조승환") {
          sendKakaoMessage(teamName, manager, "y_PK-cn9y__G6tjg1-fe79vq2fXE9sT0xfXHSw", project, task, dueDate);
        }
        if (teamName == "담당 2팀" && manager == "정종민") {
          sendKakaoMessage(teamName, manager, "y_rL_8_71-XV4NXt3-nc7MDxw_HB8MDyhQ", project, task, dueDate);
        }
        // 배정된 프로젝트 담당 매니저 일정 알림 발송
        if (manager != "조승환" && manager != "정종민") {
          sendKakaoMessage(teamName, manager, managerUUID, project, task, dueDate);
        }
      }
    }
  }
}

// 카카오톡 메시지 전송 함수
function sendKakaoMessage(team, manager, managerUUID, project, phase, dueDate) {
  // 현재 스프레드시트를 가져옵니다.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("푸시알림 설정방법");
  
  var apiUrl  = "https://kapi.kakao.com/v1/api/talk/friends/message/default/send";
  var accessToken = sheet.getRange("C8").getValue();; // 엑세스 토큰
  
  if (managerUUID == "") {
    Logger.log(`${manager}의 UUID 없음`);
    return;
  }

  var headers = {
    "Authorization": "Bearer " + accessToken,
    "Content-Type": "application/x-www-form-urlencoded;charset=utf-8"
  };

  // 템플릿에서 사용할 동적 데이터 정의
  var message = `${manager}님, '${project}' 프로젝트의 '${phase}' 단계가 ${formatDate(dueDate)}에 마감됩니다.`;

  // 템플릿에 대한 JSON 객체 구성
  var templateData = {
    "object_type": "text", // 텍스트 메시지
    "text": message, // 텍스트 메시지 내용
    "link": {
      "web_url": "https://developers.kakao.com", // 웹 링크
      "mobile_web_url": "https://developers.kakao.com" // 모바일 웹 링크
    }
  };
  
  var payload = { 
    "receiver_uuids": JSON.stringify([managerUUID]),
    "template_object": JSON.stringify(templateData),
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": true // 오류 시 예외를 무시하고 처리
  };

  Logger.log(options);

  try {
    var response = UrlFetchApp.fetch(apiUrl, options);
    Logger.log("카카오톡 메시지 응답: " + response.getContentText());
  } catch (e) {
    Logger.log("카카오톡 메시지 전송 오류: " + e.message);
  }
}

// 날짜 형식 변환 함수
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

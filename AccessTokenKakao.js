// AccessTokenKakao.gs

function getNewAccessToken() {
  // 현재 스프레드시트를 가져옵니다.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("푸시알림 설정방법");

  const tokenUrl = 'https://kauth.kakao.com/oauth/token';
  var restapi_key = sheet.getRange("C5").getValue();     // RESTAPI_KEY
  var refresh_token = sheet.getRange("C6").getValue();   // REFRESH_TOKEN
  
  const payload = {
    "grant_type": "refresh_token",
    "client_id": restapi_key,
    "refresh_token": refresh_token
  };

  const options = {
    "method": "post",
    "payload": payload,
    "muteHttpExceptions": true
  };

  try {
    var response = UrlFetchApp.fetch(tokenUrl, options);
    var jsonResponse = JSON.parse(response.getContentText());

    if (jsonResponse.error) {
      Logger.log("Access Token 갱신 오류: " + jsonResponse.error_description);
      return null;
    }

    var newAccessToken = jsonResponse.access_token;

    // Access Token이 새로 갱신되었으면 스프레드시트에 저장
    sheet.getRange("C8").setValue(newAccessToken); // Access Token 저장 (C7 셀에 저장)

    Logger.log("새 Access Token: " + newAccessToken);
    return newAccessToken;
  } catch (e) {
    Logger.log("Access Token 갱신 요청 오류: " + e.message);
    return null;
  }
}
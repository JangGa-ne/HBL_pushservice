// RefreshTokenKakao.gs

function getNewRefreshToken(e) {
  // 수동으로 실행하면 range 부분 에러남 테스트 하려면 시트에서 해야됨
  if (e.range.getSheet().getName() !== "푸시알림 설정방법") return;   // 특정 시트가 아니면 종료
  if (e.range.getA1Notation() !== "C4") return;                     // 편집된 셀이 C4가 아니면 종료

  Logger.log("C4 셀이 변경됨, 코드 실행");

  // 스프레드시트를 가져옵니다.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("푸시알림 설정방법");

  const tokenUrl = 'https://kauth.kakao.com/oauth/token';
  var oauth_code = sheet.getRange("C4").getValue();;  // OAUTH_CODE
  var restapi_key = sheet.getRange("C5").getValue();; // RESTAPI_KEY
  
  const payload = {
    'grant_type': 'authorization_code',
    'client_id': restapi_key,
    'redirect_uri': 'https://localhost',
    'code': oauth_code
  };
  
  // Authorization Code를 사용하여 Access Token과 Refresh Token 요청
  const options = {
    'method': 'post',
    'payload': payload
  };

  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const jsonResponse = JSON.parse(response.getContentText());
    
    // Access Token
    const accessToken = jsonResponse.access_token;
    const expiresIn = jsonResponse.expires_in; // 만료 시간(초 단위)
    Logger.log('Access Token: ' + accessToken);
    Logger.log('Access Token Expires In: ' + expiresIn + '초');
    
    // Refresh Token
    const refreshToken = jsonResponse.refresh_token;
    const refreshExpiresIn = jsonResponse.refresh_token_expires_in; // Refresh Token 만료 시간
    Logger.log('Refresh Token: ' + refreshToken);
    Logger.log('Refresh Token Expires In: ' + refreshExpiresIn + '초');

    // 현재 시간 (밀리초 기준)
    var now = new Date();
    // Refresh Token 만료 시간 계산 (현재 시간 + expiresIn 초)
    var refreshTokenExpiryDate = new Date(now.getTime() + refreshExpiresIn * 1000);

    // refresh_token 등록
    sheet.getRange("C6").setValue(refreshToken);
    sheet.getRange("C7").setValue(refreshTokenExpiryDate);
  } catch (e) {
    Logger.log('Error: ' + e.message);
  }
}
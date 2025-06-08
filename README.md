<h1>🔐 <code>RefreshTokenKakao.gs</code> 스크립트 설명</h1>

  <p>
    이 Google Apps Script는 카카오 OAuth 인증을 통해 사용자의 <strong>Authorization Code</strong>로부터
    <strong>Access Token</strong>과 <strong>Refresh Token</strong>을 자동으로 요청하고,
    이를 Google Sheet에 기록하는 자동화 스크립트입니다.
  </p>

  <h2>📌 트리거 조건</h2>
  <p>
    아래 조건이 모두 충족될 때 스크립트가 실행됩니다:
  </p>
  <ul>
    <li>시트 이름이 <strong>"푸시알림 설정방법"</strong>일 것</li>
    <li>셀 <code>C4</code>가 편집되었을 것</li>
  </ul>

  <pre><code>if (e.range.getSheet().getName() !== "푸시알림 설정방법") return;
if (e.range.getA1Notation() !== "C4") return;</code></pre>

  <h2>🔄 주요 흐름 요약</h2>
  <ol>
    <li><strong>Authorization Code</strong>를 셀 <code>C4</code>에서 읽어옴</li>
    <li><strong>REST API Key</strong>를 셀 <code>C5</code>에서 읽어옴</li>
    <li>카카오 토큰 API (<code>https://kauth.kakao.com/oauth/token</code>)로 POST 요청</li>
    <li>응답에서 Access Token과 Refresh Token을 파싱</li>
    <li>
      Refresh Token과 만료 일시를 아래 셀에 기록:
      <ul>
        <li><code>C6</code>: Refresh Token</li>
        <li><code>C7</code>: Refresh Token 만료 일시</li>
      </ul>
    </li>
  </ol>

  <h2>📤 카카오 토큰 요청</h2>
  <pre><code>{
  grant_type: 'authorization_code',
  client_id: [REST API 키],
  redirect_uri: 'https://localhost',
  code: [Authorization Code]
}</code></pre>

  <h2>📥 응답 예시 (JSON)</h2>
  <pre><code>{
  "access_token": "액세스토큰",
  "token_type": "bearer",
  "refresh_token": "리프레시토큰",
  "expires_in": 21599,
  "refresh_token_expires_in": 5184000
}</code></pre>

  <h2>📝 결과 저장</h2>
  <ul>
    <li><strong>Access Token</strong>은 로그에만 기록됨 (<code>Logger.log</code>)</li>
    <li><strong>Refresh Token</strong>은 셀 <code>C6</code>에 저장</li>
    <li><strong>만료 시간</strong>은 현재 시각 기준으로 계산하여 <code>C7</code>에 저장</li>
  </ul>

  <h2>⚠️ 테스트 주의사항</h2>
  <p>
    <strong>이벤트 객체 <code>e</code>가 포함된 함수이므로, 수동 실행 시 <code>e</code>가 정의되지 않아 오류 발생</strong>합니다.
    반드시 <code>스프레드시트에서 C4 셀을 편집</code>함으로써 자동 실행되어야 정상 작동합니다.
  </p>

  <h2>🚨 예외 처리</h2>
  <p>
    <code>UrlFetchApp.fetch()</code> 과정에서 오류 발생 시 <code>Logger.log</code>로 오류 메시지를 출력합니다.
  </p>

  <h2>📚 참고</h2>
  <ul>
    <li><a href="https://developers.kakao.com/docs/latest/ko/kakaologin/rest-api#request-token" target="_blank">카카오 REST API 문서: 토큰 요청</a></li>
    <li><a href="https://developers.google.com/apps-script/guides/triggers" target="_blank">Google Apps Script 트리거 문서</a></li>
  </ul>

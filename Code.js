// Code.gs

function createTimeDrivenTriggers() {
  // 기존 트리거 삭제 (중복 방지)
  deleteExistingTriggers("sendPush");
  // 매일 오후 5시에 실행되는 트리거 생성
  ScriptApp.newTrigger('sendPush')
    .timeBased()
    .atHour(8)         // 8-9시 실행
    .everyDays(1)      // 매일 실행
    .create();
}

function deleteExistingTriggers(functionName) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function sendPush() {
  // 카카오톡으로 알림 발송
  sendKakao();
}
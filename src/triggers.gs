/**
 * 実行トリガーの定義
 * スケジュールを変えたら、clasp push の後にエディタで setupTriggers を1回実行する
 * （push だけではトリガーは変わらない）
 */

const TRIGGER_HANDLER = 'processEmailToSlack';

/**
 * processEmailToSlack を毎週水曜の9時台（appsscript.json の timeZone 基準）に実行するトリガーを作る
 * 何度実行しても重複しないよう、同じ関数の既存トリガーを消してから作り直す
 */
function setupTriggers() {
  ScriptApp.getProjectTriggers()
    .filter((trigger) => trigger.getHandlerFunction() === TRIGGER_HANDLER)
    .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger(TRIGGER_HANDLER)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.WEDNESDAY)
    .atHour(9)
    .create();

  Logger.log(`${TRIGGER_HANDLER} のトリガーを作り直しました（毎週水曜 9時台）`);
}

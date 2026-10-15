const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// トリガーの作成・削除を記録する偽の ScriptApp
const createFakeScriptApp = (existingHandlers) => {
  const triggers = existingHandlers.map((handler) => ({ getHandlerFunction: () => handler }));
  const fake = {
    deleted: [],
    created: [],
    WeekDay: { WEDNESDAY: 'WEDNESDAY' },
    getProjectTriggers: () => triggers,
    deleteTrigger: (trigger) => fake.deleted.push(trigger.getHandlerFunction()),
    newTrigger: (handler) => {
      const spec = { handler };
      const builder = {
        timeBased: () => builder,
        onWeekDay: (day) => { spec.weekDay = day; return builder; },
        atHour: (hour) => { spec.hour = hour; return builder; },
        create: () => { fake.created.push(spec); },
      };
      return builder;
    },
  };
  return fake;
};

const loadTriggers = (ScriptApp) => {
  const context = { ScriptApp, Logger: { log: () => {} } };
  vm.createContext(context);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'src', 'triggers.gs'), 'utf8'), context);
  return context;
};

test('processEmailToSlack を毎週水曜9時台に実行するトリガーを作る', () => {
  const ScriptApp = createFakeScriptApp([]);

  vm.runInContext('setupTriggers()', loadTriggers(ScriptApp));

  assert.deepEqual(ScriptApp.created, [{ handler: 'processEmailToSlack', weekDay: 'WEDNESDAY', hour: 9 }]);
});

test('既存の processEmailToSlack のトリガーだけを消してから作り直す', () => {
  const ScriptApp = createFakeScriptApp(['processEmailToSlack', 'otherFunction', 'processEmailToSlack']);

  vm.runInContext('setupTriggers()', loadTriggers(ScriptApp));

  assert.deepEqual(ScriptApp.deleted, ['processEmailToSlack', 'processEmailToSlack']);
  assert.equal(ScriptApp.created.length, 1);
});

#!/usr/bin/env node
// Скрипт «tick» для каталога банков. У ЦБ нет публичного XLSX-эндпоинта
// со списком банков (страница /banking_sector/credit/ работает через JS).
// Поэтому каталог поддерживается вручную через PR, а скрипт:
//
//   1. Валидирует JSON (структуру, обязательные поля).
//   2. Делает легковесный sanity-check через XML_bic.asp:
//      печатает, сколько банков из нашего seed нашлось в публичном
//      справочнике BIC ЦБ. На основании этого редактор может вручную
//      пометить отзывы лицензий.
//   3. При наличии изменений — поднимает version до сегодняшней даты
//      (с .N-суффиксом, если уже был запуск сегодня).
//
// Запуск:
//   node scripts/update.mjs

import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CATALOG_PATH = path.join(__dirname, '..', 'creditors.json');

const CBR_BIC_URL = 'http://www.cbr.ru/scripts/XML_bic.asp';

const todayIso = () => new Date().toISOString();
const todayVersion = () => {
  const d = new Date();
  return `${d.getUTCFullYear()}.${String(d.getUTCMonth() + 1).padStart(2, '0')}.${String(d.getUTCDate()).padStart(2, '0')}`;
};

async function fetchBicSet() {
  try {
    const res = await fetch(CBR_BIC_URL);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    const text = new TextDecoder('windows-1251').decode(buf);
    const set = new Set();
    const re = /<Bic>([^<]+)<\/Bic>/g;
    let m;
    while ((m = re.exec(text)) !== null) set.add(m[1].trim());
    return set;
  } catch (e) {
    console.warn(`⚠ Не удалось скачать BIC справочник: ${e.message ?? e}`);
    return null;
  }
}

function validate(data) {
  const errors = [];
  if (!data.version) errors.push('отсутствует поле version');
  if (!Array.isArray(data.creditors)) errors.push('creditors не массив');
  for (const [i, c] of (data.creditors ?? []).entries()) {
    if (!c.inn) errors.push(`#${i}: нет inn`);
    if (!c.name) errors.push(`#${i}: нет name`);
    if (!['active', 'suspended', 'revoked', 'unknown'].includes(c.licenseStatus)) {
      errors.push(`#${i} (${c.name}): неизвестный licenseStatus "${c.licenseStatus}"`);
    }
  }
  return errors;
}

async function readCurrent() {
  const raw = await fs.readFile(CATALOG_PATH, 'utf-8');
  return JSON.parse(raw);
}

async function main() {
  const current = await readCurrent();
  console.log(`Текущая версия: ${current.version} (${current.creditors.length} записей)`);

  // 1) Валидация
  const errors = validate(current);
  if (errors.length) {
    console.error('✗ Каталог невалиден:');
    for (const e of errors) console.error('   -', e);
    process.exit(1);
  }
  console.log('✓ JSON валиден');

  // 2) Sanity-check через XML_bic.asp (не отзываем автоматически — только репорт)
  const bicSet = await fetchBicSet();
  if (bicSet) {
    let confirmed = 0;
    let unconfirmed = 0;
    const unconfirmedNames = [];
    for (const c of current.creditors) {
      if (c.inn === '0000000000') continue;
      if (!c.bik) { unconfirmed++; continue; }
      if (bicSet.has(c.bik)) confirmed++;
      else {
        unconfirmed++;
        if (c.licenseStatus === 'active') {
          unconfirmedNames.push(`${c.shortName ?? c.name} (BIC ${c.bik})`);
        }
      }
    }
    console.log(`  активных подтверждено в BIC-выборке: ${confirmed}`);
    console.log(`  не нашли в публичной выборке:        ${unconfirmed}`);
    if (unconfirmedNames.length) {
      console.log('  ⓘ XML_bic.asp выдаёт ограниченную территориальную выборку,');
      console.log('    отсутствие BIC в ней НЕ означает отзыв лицензии.');
      console.log('    Проверьте при необходимости вручную:');
      for (const n of unconfirmedNames.slice(0, 10)) console.log('     -', n);
      if (unconfirmedNames.length > 10) {
        console.log(`     … и ещё ${unconfirmedNames.length - 10}`);
      }
    }
  }

  // 3) Bump версии (новая дата → новая версия). Редакторские правки в JSON
  // подхватываются именно так: после ручного PR на main воркфлоу запускает
  // скрипт, и version обновляется до текущей даты.
  const today = todayVersion();
  let nextVersion = today;
  if (current.version === today) {
    nextVersion = `${today}.1`;
  } else if (current.version.startsWith(`${today}.`)) {
    const m = /\.(\d+)$/.exec(current.version);
    const n = m ? Number(m[1]) + 1 : 1;
    nextVersion = `${today}.${n}`;
  }

  const next = {
    ...current,
    version: nextVersion,
    updatedAt: todayIso(),
  };

  // Записываем только если что-то изменилось
  const before = JSON.stringify(current);
  const after = JSON.stringify(next);
  if (before === after) {
    console.log('  Изменений нет — пропускаем запись.');
    return;
  }

  await fs.writeFile(
    CATALOG_PATH,
    JSON.stringify(next, null, 2) + '\n',
    'utf-8',
  );
  console.log('✓ Каталог обновлён');
  console.log(`  Версия: ${current.version} → ${nextVersion}`);
}

main().catch((e) => {
  console.error('✗ Ошибка:', e.message ?? e);
  process.exit(1);
});

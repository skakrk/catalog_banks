#!/usr/bin/env node
// Обновление каталога банков из реестра ЦБ РФ.
//
// Источник: страница «Информация по кредитным организациям» (FullCoList),
// которая внутри имеет кнопку экспорта в Excel:
//   https://www.cbr.ru/Queries/UniDbQuery/DownloadExcel/98547
//   ?FromDate=DD/MM/YYYY&ToDate=DD/MM/YYYY&posted=False
//
// Лист RC содержит ~1900 кредитных организаций (действующие + отозванные +
// аннулированные + ликвидированные). В наш каталог импортируем только
// Действующие и Отозванные — этого достаточно для пользователя.
//
// Колонки реестра:
//   rn         — порядковый номер
//   bnk_type   — тип (НКО, расчётная и т.п.; пустой = универсальный банк)
//   cregnum    — регистрационный номер ЦБ (= лицензия)
//   ogrn       — ОГРН (число в Excel)
//   bnk_name   — наименование
//   opf        — организационно-правовая форма
//   reg_date   — дата регистрации (Excel-число дней)
//   lic_status — Действующая | Отозванная | Аннулированная | Ликвидация
//   bnk_addr   — адрес
//
// В реестре нет ИНН, поэтому для банков, которых ещё нет в каталоге,
// ставим псевдо-ИНН вида `cbr-<cregnum>`. Когда наполнение делается через
// PR — редактор может вручную дополнить настоящим ИНН и БИКом.

import fs from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

let XLSX;
try {
  XLSX = (await import('xlsx')).default;
} catch {
  console.error('Нужен пакет xlsx. Запустите: npm install');
  process.exit(1);
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CATALOG_PATH = path.join(__dirname, '..', 'creditors.json');

const formatCbrDate = (d) => {
  // ЦБ ожидает MM/DD/YYYY (US-style) — проверено по странице FullCoList,
  // hidden-форма передаёт именно такой формат.
  const mm = String(d.getUTCMonth() + 1).padStart(2, '0');
  const dd = String(d.getUTCDate()).padStart(2, '0');
  return `${mm}/${dd}/${d.getUTCFullYear()}`;
};

const buildCbrUrl = () => {
  const today = formatCbrDate(new Date());
  const params = new URLSearchParams({
    FromDate: today,
    ToDate: today,
    posted: 'False',
  });
  return `https://www.cbr.ru/Queries/UniDbQuery/DownloadExcel/98547?${params.toString()}`;
};

const todayIso = () => new Date().toISOString();
const todayVersion = () => {
  const d = new Date();
  return `${d.getUTCFullYear()}.${String(d.getUTCMonth() + 1).padStart(2, '0')}.${String(d.getUTCDate()).padStart(2, '0')}`;
};

async function downloadXlsx() {
  const url = buildCbrUrl();
  console.log(`→ Скачиваю реестр: ${url}`);
  const res = await fetch(url, { headers: { 'User-Agent': 'Mozilla/5.0 catalog_banks updater' } });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  console.log(`  получено ${buf.length} байт`);
  return buf;
}

const STATUS_MAP = {
  'Действующая':   'active',
  'Отозванная':    'revoked',
  'Аннулированная':'revoked',
  'Ликвидация':    'revoked',
};

function parseRegistry(buf) {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const sheet = wb.Sheets['RC'] ?? wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });

  const out = [];
  for (const r of rows) {
    const status = STATUS_MAP[r.lic_status];
    if (!status) continue;                          // пустой/неизвестный статус — пропускаем
    if (status !== 'active' && r.lic_status !== 'Отозванная') {
      // Импортируем только Действующие и Отозванные. Аннулированные/
      // Ликвидированные тоже становятся 'revoked' выше — но мы оставим
      // только тех, у кого был полноценный отзыв (Отозванная). Иначе
      // каталог раздуется на 1500 «исторических» записей.
      continue;
    }
    const cregnum = r.cregnum != null ? String(r.cregnum).trim() : '';
    if (!cregnum) continue;
    out.push({
      cregnum,
      name: String(r.bnk_name ?? '').trim(),
      ogrn: r.ogrn != null ? String(r.ogrn).trim() : undefined,
      address: r.bnk_addr ? String(r.bnk_addr).trim() : undefined,
      bnk_type: r.bnk_type ? String(r.bnk_type).trim() : undefined,
      licenseStatus: status,
    });
  }
  return out;
}

async function readCurrent() {
  const raw = await fs.readFile(CATALOG_PATH, 'utf-8');
  return JSON.parse(raw);
}

const PSEUDO_INN_PREFIX = 'cbr-';

function merge(current, fromRegistry) {
  // Индекс существующих записей по cregnum (= licenseNo) и по inn
  const byCregnum = new Map();
  for (const c of current.creditors) {
    if (c.licenseNo) byCregnum.set(String(c.licenseNo), c);
  }

  const now = todayIso();
  const seenCregnums = new Set();
  let added = 0;
  let updated = 0;

  for (const r of fromRegistry) {
    seenCregnums.add(r.cregnum);
    const existing = byCregnum.get(r.cregnum);
    if (existing) {
      // Обновляем поля из реестра, не трогая редакторские (ИНН, БИК,
      // ratingScore, website, shortName).
      existing.name = r.name || existing.name;
      existing.ogrn = r.ogrn ?? existing.ogrn;
      existing.address = r.address ?? existing.address;
      const wasRevoked = existing.licenseStatus === 'revoked';
      existing.licenseStatus = r.licenseStatus;
      if (r.licenseStatus === 'revoked' && !wasRevoked) {
        existing.revokedAt = now.slice(0, 10);
      } else if (r.licenseStatus === 'active') {
        delete existing.revokedAt;
      }
      existing.updatedAt = now;
      updated++;
    } else {
      // Новая запись — без ИНН и БИК, с псевдо-ИНН для совместимости с
      // приложением (которое использует ИНН как primary key).
      const newRec = {
        inn: `${PSEUDO_INN_PREFIX}${r.cregnum}`,
        name: r.name,
        ogrn: r.ogrn,
        licenseNo: r.cregnum,
        registryRecordNo: r.cregnum,
        licenseStatus: r.licenseStatus,
        address: r.address,
        complaintsUrl: 'https://www.cbr.ru/Reception/Message/Register',
        ratingScore: 50,
        updatedAt: now,
      };
      if (r.licenseStatus === 'revoked') {
        newRec.revokedAt = now.slice(0, 10);
      }
      current.creditors.push(newRec);
      added++;
    }
  }

  // Записи в каталоге, которых нет в реестре, но есть cregnum — отзыв
  // лицензии на стороне ЦБ может означать, что банк выпал из выдачи.
  // Не трогаем — оставляем последний известный статус.

  // Сортировка: плейсхолдер первым, затем active, затем revoked, по name
  current.creditors.sort((a, b) => {
    if (a.inn === '0000000000') return -1;
    if (b.inn === '0000000000') return 1;
    const aActive = a.licenseStatus === 'active' ? 0 : 1;
    const bActive = b.licenseStatus === 'active' ? 0 : 1;
    if (aActive !== bActive) return aActive - bActive;
    return a.name.localeCompare(b.name, 'ru');
  });

  // Версия с авто-инкрементом за день
  const today = todayVersion();
  let nextVersion = today;
  if (current.version === today) {
    nextVersion = `${today}.1`;
  } else if (current.version.startsWith(`${today}.`)) {
    const m = /\.(\d+)$/.exec(current.version);
    const n = m ? Number(m[1]) + 1 : 1;
    nextVersion = `${today}.${n}`;
  }

  return {
    next: {
      ...current,
      version: nextVersion,
      updatedAt: now,
    },
    stats: {
      total: current.creditors.length,
      active: current.creditors.filter((c) => c.licenseStatus === 'active').length,
      revoked: current.creditors.filter((c) => c.licenseStatus === 'revoked').length,
      added,
      updated,
    },
  };
}

async function main() {
  const current = await readCurrent();
  console.log(`Текущая версия: ${current.version} (${current.creditors.length} записей)`);

  const buf = await downloadXlsx();
  const fromReg = parseRegistry(buf);
  console.log(`→ В реестре: ${fromReg.length} банков (active + revoked)`);

  const { next, stats } = merge(current, fromReg);

  await fs.writeFile(
    CATALOG_PATH,
    JSON.stringify(next, null, 2) + '\n',
    'utf-8',
  );

  console.log('✓ Каталог обновлён');
  console.log(`  Версия: ${next.version}`);
  console.log(`  Всего записей: ${stats.total} (новых: ${stats.added}, обновлено: ${stats.updated})`);
  console.log(`    действующих:  ${stats.active}`);
  console.log(`    отозванных:   ${stats.revoked}`);
}

main().catch((e) => {
  console.error('✗ Ошибка:', e.message ?? e);
  process.exit(1);
});

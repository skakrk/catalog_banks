# catalog_banks

Открытый JSON-справочник банков РФ. Базируется на топ-50 действующих банков
с ручной редактурой и автоматически обогащается актуальными короткими
названиями из реестра BIC ЦБ РФ.

## Структура

- `creditors.json` — текущий каталог банков
- `scripts/update.mjs` — обновление через XML_bic.asp (CBR BIC registry)
- `package.json` — зависимости скрипта обновления

## Формат `creditors.json`

```jsonc
{
  "version": "YYYY.MM.DD",
  "updatedAt": "ISO 8601",
  "source": "https://www.cbr.ru/banking_sector/credit/",
  "creditors": [
    {
      "inn": "0000000000",
      "name": "...",
      "shortName": "...",
      "ogrn": "...",
      "licenseNo": "...",
      "licenseStatus": "active | suspended | revoked | unknown",
      "bik": "044525225",
      "updatedAt": "ISO 8601"
    }
  ]
}
```

## Источник данных

- Базовый список и реквизиты — реестр кредитных организаций ЦБ РФ:
  https://www.cbr.ru/banking_sector/credit/
- Обновление shortName / признаков активности — справочник BIC:
  http://www.cbr.ru/scripts/XML_bic.asp

## Лицензия

Данные ЦБ РФ — общедоступная информация.
JSON-структура и скрипт — MIT.

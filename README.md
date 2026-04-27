# Generátor Ganttova diagramu pro Excel

> Profesionální Excel šablona Ganttova diagramu — zdarma, open-source, MIT licence.
> **Live tool: <https://autoerp.cz/ganttuv-diagram-generator>**

Hotová šablona harmonogramu projektu pro Excel, LibreOffice Calc a Google Sheets. Stáhněte si `.xlsx`, otevřete a začněte plánovat. Přebírejte termíny, sledujte pokrok a tiskněte na A4 — bez registrace.

## Co je uvnitř

`ganttuv-diagram-vzor.xlsx` obsahuje 4 listy:

1. **Ganttův diagram** — hlavní pohled. Tabulka úkolů (ID, název, vlastník, termíny, status, závislost) + 26týdenní časová osa s automatickým vykreslováním pruhů přes podmíněné formátování. 17 vzorových úkolů ve fázích Příprava → Analýza → Implementace → Testování → Spuštění (projekt „Implementace ERP").
2. **Vizualizace** — denní rozlišení Ganttovy vizualizace s pruhy plánováno (modrá) a hotová část (žlutá).
3. **Návod** — 8 kroků, jak šablonu používat.
4. **O šabloně** — licence, zdrojový kód, AutoERP CTA.

### Funkce

- Automatický výpočet délky úkolu (`=Konec−Začátek+1`)
- Vážený pokrok celého projektu
- Datová validace ve sloupci Status (Nezahájeno / Probíhá / Hotovo / Zpožděno) s barevným označením
- Color-scale podmíněné formátování pro pokrok 0 % → 50 % → 100 %
- Časová osa po týdnech, automaticky se vyplňuje na základě termínů úkolu a aktuálního pokroku
- A4 na šířku, fit na 1 stránku, připraveno k tisku

## Stažení

Nejnovější release jako .xlsx:

➜ **<https://github.com/Draivix/ganttuv-diagram-generator/releases/latest>**

Nebo přímo z webu:

➜ **<https://autoerp.cz/templates/ganttuv-diagram-vzor.xlsx>**

## Otevřít issue / nahlásit chybu

Našli jste chybu, vzorec nefunguje, formátování se rozbilo? Otevřete issue:

➜ <https://github.com/Draivix/ganttuv-diagram-generator/issues/new>

Reagujeme zpravidla do 24 hodin.

## Generování ze zdroje

Šablona je vygenerovaná ze skriptu `scripts/generate.ts` (TypeScript, ESM). Generátor používá `exceljs` pro vytvoření plně funkční šablony se vzorci, podmíněným formátováním a datovými validacemi.

```bash
npm install exceljs tsx
npx tsx scripts/generate.ts
# vygeneruje ganttuv-diagram-vzor.xlsx
```

## Licence

MIT — viz [LICENSE](LICENSE). Můžete šablonu volně používat, upravovat a šířit i komerčně.

## Kdo to dělá

[**AutoERP**](https://autoerp.cz) — modulární ERP/CRM pro české a slovenské firmy.
Provozovatel: Apertia Tech s.r.o. (IČO 27117758, Praha).

Pokud řídíte víc než jeden projekt zároveň, máte tým 5+ lidí nebo chcete kapacitní plánování zdrojů, podívejte se na [modul Projektové řízení v AutoERP](https://autoerp.cz/projektove-rizeni). Od 3 450 Kč/měsíc, bez licencí za uživatele, 14 dní zdarma.

---

## English

A free Gantt-chart Excel template + interactive web tool. Same workbook, English notes below.

The `.xlsx` works in Microsoft Excel (2016+), LibreOffice Calc (6+), and Google Sheets. Four sheets: main Gantt view with formula-driven 26-week timeline, daily-resolution visualization, instructions, and an "About" sheet. All formulas (duration, weighted progress, conditional bar fill) use cross-engine standards.

- **Live tool:** <https://autoerp.cz/ganttuv-diagram-generator>
- **Download .xlsx:** <https://github.com/Draivix/ganttuv-diagram-generator/releases/latest>
- **License:** MIT
- **Maintainer:** [Apertia Tech s.r.o.](https://autoerp.cz) — makers of AutoERP
- **Issues / bugs:** <https://github.com/Draivix/ganttuv-diagram-generator/issues>

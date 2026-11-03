import { createUniver, LocaleType, mergeLocales, greenTheme, CommandType } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import CoreEn from '@univerjs/preset-sheets-core/locales/en-US';
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import FilterEn from '@univerjs/preset-sheets-filter/locales/en-US';
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import SortEn from '@univerjs/preset-sheets-sort/locales/en-US';
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import FindEn from '@univerjs/preset-sheets-find-replace/locales/en-US';
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation';
import ValidationEn from '@univerjs/preset-sheets-data-validation/locales/en-US';
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting';
import ConditionalEn from '@univerjs/preset-sheets-conditional-formatting/locales/en-US';
import '@univerjs/preset-sheets-core/lib/index.css';
import '@univerjs/preset-sheets-filter/lib/index.css';
import '@univerjs/preset-sheets-sort/lib/index.css';
import '@univerjs/preset-sheets-find-replace/lib/index.css';
import '@univerjs/preset-sheets-data-validation/lib/index.css';
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css';
import { BookRecord, Workbook } from './model';

export function createEngine(container: HTMLElement, book: BookRecord, changed: () => void, selected: (address: string, raw: string, displayed: string) => void, dark: boolean) {
  const { univer, univerAPI: api } = createUniver({
    locale: LocaleType.EN_US,
    locales: { [LocaleType.EN_US]: mergeLocales(CoreEn, FilterEn, SortEn, FindEn, ValidationEn, ConditionalEn) },
    theme: greenTheme, darkMode: dark,
    presets: [UniverSheetsCorePreset({ container, header: true, formulaBar: false, toolbar: true, ribbonType: 'classic' }), UniverSheetsFilterPreset(), UniverSheetsSortPreset(), UniverSheetsFindReplacePreset(), UniverSheetsDataValidationPreset(), UniverSheetsConditionalFormattingPreset()],
  });
  api.createWorkbook(book.workbook as any);
  const refresh = () => { const range = api.getActiveWorkbook()?.getActiveSheet().getActiveRange(); if (range) selected(range.getA1Notation(), range.getFormula() || String(range.getValue() ?? ''), range.getDisplayValue()); };
  const commandListener = api.addEvent(api.Event.CommandExecuted, event => { if (event.type === CommandType.MUTATION) changed(); if (!event.id.includes('scroll')) refresh(); });
  const selectionListener = api.addEvent(api.Event.SelectionChanged, refresh);
  refresh();
  return { api, save: () => structuredClone(api.getActiveWorkbook()!.save()) as unknown as Workbook, dispose: () => { commandListener.dispose(); selectionListener.dispose(); univer.dispose(); } };
}
export type SheetEngine = ReturnType<typeof createEngine>;

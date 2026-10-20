import { it, expect } from 'vitest';
import { loadApp } from './harness.js';

it('decodes UTF-16 BOM exports', () => {
  const { run } = loadApp();
  expect(run('decodeText(new Uint8Array([255,254,97,0,44,0,98,0]).buffer)')).toBe('a,b');
});
it('converts a workbook to named CSVs', () => {
  const { run } = loadApp();
  const result = run('workbookToCsvs("report.xlsx", {SheetNames:["Data"],Sheets:{Data:XLSX.utils.aoa_to_sheet([["name","value"],["Ada",42]])}})');
  expect(result[0].name).toBe('report.csv');
  expect(result[0].csv).toContain('Ada,42');
});
it('starts with downloading disabled', () => {
  expect(loadApp().window.document.getElementById('run').disabled).toBe(true);
});

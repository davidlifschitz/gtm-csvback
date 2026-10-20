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

it('quotes commas and preserves quoted tabs when converting TSV', async () => {
  const { run } = loadApp();
  const input = 'name\tnote\nAda\t"hello, world"\nBob\t"left\tright"\n';
  const encoded = Array.from(new TextEncoder().encode(input));
  const result = await run(`convertFile({name:"input.tsv",type:"text/tab-separated-values",arrayBuffer:async()=>new Uint8Array(${JSON.stringify(encoded)}).buffer})`);
  expect(result[0].csv).toContain('Ada,"hello, world"');
  expect(result[0].csv).toContain('Bob,left\tright');
});
it('uses the same TSV conversion for archive tables', async () => {
  const { run } = loadApp();
  const result = await run('(async()=>{const z=new JSZip();z.file("table.tsv","name\\tnote\\nAda\\thello, world");return fromNumbersZip("book.numbers",await z.generateAsync({type:"uint8array"}));})()');
  expect(result[0].csv).toContain('Ada,"hello, world"');
});

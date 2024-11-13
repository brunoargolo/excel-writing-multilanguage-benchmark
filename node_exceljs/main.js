import ExcelJS from 'exceljs';
import { readFile } from 'fs/promises';
import { gunzip } from 'zlib';
import { promisify } from 'util';

const gunzipAsync = promisify(gunzip);

async function readCompressedJsonFile(filename) {
    const compressedData = await readFile(filename);
    const jsonString = await gunzipAsync(compressedData);
    return JSON.parse(jsonString);
}

async function writeToExcel(records) {
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'demo.xlsx' });
    
    // Get N_SHEETS value from environment variable, default to 1 if not set
    const nSheets = parseInt(process.env.N_SHEETS) || 1;
    
    // Validate N_SHEETS value
    if (nSheets < 1 || nSheets > 9) {
        console.error('N_SHEETS must be between 1 and 9');
        process.exit(1);
    }

    const columns = [
        { header: 'ID', key: 'id', width: 22 },
        { header: 'My String 1', key: 'myString1', width: 22 },
        { header: 'My Numeric String', key: 'myNumericString', width: 22 },
        { header: 'My Strign 2', key: 'myString2', width: 22 },
        { header: 'Amount', key: 'amount', width: 15, style: { numFmt: '0.000' } },
        { header: 'My Date 1', key: 'myDate1', width: 15, style: { numFmt: 'yyyy-mm-dd' } },
        { header: 'My Date 2', key: 'myDate2', width: 15, style: { numFmt: 'yyyy-mm-dd' } }
    ];

    const promises = [];
    // Create N sheets
    for (let i = 1; i <= nSheets; i++) {
        promises.push(createSheet(workbook, i, columns, records));
    }
    await Promise.all(promises);

    await workbook.commit();
}

async function createSheet(workbook, i, columns, records) {
    const worksheet = workbook.addWorksheet(`Sheet${i}`);
    worksheet.columns = columns;

    for (const record of records) {
        worksheet.addRow(record).commit();
    }
    await worksheet.commit();
}

async function main() {
    console.time('Load Time');
    const records = await readCompressedJsonFile('../input.json.gzip');
    console.timeEnd('Load Time');

    console.log(`Retrieved ${records.length} records`);

    console.time('Write Time');
    await writeToExcel(records);
    console.timeEnd('Write Time');
}

await main();
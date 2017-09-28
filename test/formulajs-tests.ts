import * as formulajs from 'formulajs';
import * as assert from 'assert';
import XLSX_CALC, { import_functions } from '../src';

describe('formulajs integration', function() {
    describe('XLSX_CALC.import_functions()', function() {
        it('imports the functions from formulajs', function() {
            import_functions(formulajs);
            const workbook: any = {};
            workbook.Sheets = {};
            workbook.Sheets.Sheet1 = {};
            workbook.Sheets.Sheet1.A1 = {v: 2};
            workbook.Sheets.Sheet1.A2 = {v: 4};
            workbook.Sheets.Sheet1.A3 = {v: 8};
            workbook.Sheets.Sheet1.A4 = {v: 16};
            workbook.Sheets.Sheet1.A5 = {f: 'AVERAGEIF(A1:A4,">5")'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A5.v, 12);
        });
        it('imports the functions with dot names like BETA.DIST', function() {
            import_functions(formulajs);
            const workbook: any = {Sheets: {Sheet1: {}}};
            workbook.Sheets.Sheet1.A5 = {f: 'BETA.DIST(2, 8, 10, true, 1, 3)'};
            XLSX_CALC(workbook);
            assert.equal(workbook.Sheets.Sheet1.A5.v.toFixed(10), (0.6854705810117458).toFixed(10));
        });
    });
});

import { test } from 'node:test';
import assert from 'node:assert/strict';

import { __setGraphClientInstanceForTests } from '../graph/client.js';
import { handleExcelAnalysis, handleExcelOperations } from '../tools/advanced/excel.js';
import { registerGraphClientTestLifecycle } from './helpers/test-lifecycle.js';
import { createMockGraphClient, parsePayload, type ToolEnvelope } from './helpers/tool-test-helpers.js';

registerGraphClientTestLifecycle();

test('excel_operations resolves itemPath, creates session, executes operation, and closes session', async () => {
  const mock = createMockGraphClient({
    get: async (endpoint: string, params?: Record<string, unknown>, options?: Record<string, unknown>) => {
      if (endpoint === '/sites/site-1/drive/root:/Reports/book.xlsx') {
        return { success: true, data: { id: 'excel-1' } };
      }

      if (endpoint === '/sites/site-1/drive/items/excel-1/workbook/worksheets/SheetA/range(address=\'A1:B2\')') {
        assert.deepEqual(params, {});
        assert.deepEqual(options, { headers: { 'workbook-session-id': 'session-1' } });
        return {
          success: true,
          data: {
            address: 'SheetA!A1:B2',
            rowCount: 2,
            columnCount: 2,
            values: [[1, 2], [3, 4]],
            formulas: [[null, null], [null, null]],
            text: [['1', '2'], ['3', '4']],
            numberFormat: [['General', 'General'], ['General', 'General']],
          },
        };
      }

      throw new Error(`Unexpected GET ${endpoint}`);
    },
    post: async (endpoint: string, data?: unknown, options?: Record<string, unknown>) => {
      if (endpoint.endsWith('/createSession')) {
        assert.deepEqual(data, { persistChanges: true });
        return { success: true, data: { id: 'session-1' } };
      }

      if (endpoint.endsWith('/closeSession')) {
        assert.deepEqual(options, { headers: { 'workbook-session-id': 'session-1' } });
        return { success: true };
      }

      throw new Error(`Unexpected POST ${endpoint}`);
    },
  });
  __setGraphClientInstanceForTests(mock.client as any);

  const response = (await handleExcelOperations({
    itemPath: 'Reports/book.xlsx',
    siteId: 'site-1',
    operation: 'read_range',
    worksheet: 'SheetA',
    range: 'A1:B2',
    useSession: true,
  })) as ToolEnvelope;
  const payload = parsePayload<any>(response);

  assert.equal(response.isError, undefined);
  assert.equal(payload.operation, 'read_range');
  assert.equal(payload.address, 'SheetA!A1:B2');
  assert.deepEqual(payload.values, [[1, 2], [3, 4]]);
  assert.deepEqual(
    mock.methodCalls('post').map((call) => call.args[0]),
    [
      '/sites/site-1/drive/items/excel-1/workbook/createSession',
      '/sites/site-1/drive/items/excel-1/workbook/closeSession',
    ],
  );
});

test('excel_operations covers workbook operation branches', async () => {
  const scenarios = [
    {
      name: 'write_range',
      args: { itemId: 'excel-1', operation: 'write_range', worksheet: 'Sheet1', range: 'A1:B2', values: [[1, 2], [3, 4]] },
      method: 'patch',
      endpoint: '/me/drive/items/excel-1/workbook/worksheets/Sheet1/range(address=\'A1:B2\')',
      resultCheck: (payload: any) => {
        assert.equal(payload.rowsWritten, 2);
        assert.equal(payload.columnsWritten, 2);
        assert.equal(payload.address, 'Sheet1!A1:B2');
      },
    },
    {
      name: 'add_worksheet',
      args: { itemId: 'excel-1', operation: 'add_worksheet', worksheet: 'Backlog' },
      method: 'post',
      endpoint: '/me/drive/items/excel-1/workbook/worksheets/add',
      resultCheck: (payload: any) => {
        assert.equal(payload.worksheet.name, 'Backlog');
        assert.equal(payload.worksheet.position, 1);
      },
    },
    {
      name: 'list_worksheets',
      args: { itemId: 'excel-1', operation: 'list_worksheets' },
      method: 'get',
      endpoint: '/me/drive/items/excel-1/workbook/worksheets',
      resultCheck: (payload: any) => {
        assert.equal(payload.count, 2);
        assert.equal(payload.worksheets[1].name, 'Backlog');
      },
    },
    {
      name: 'get_formulas',
      args: { itemId: 'excel-1', operation: 'get_formulas', worksheet: 'Sheet1', range: 'C1:C2' },
      method: 'get',
      endpoint: '/me/drive/items/excel-1/workbook/worksheets/Sheet1/range(address=\'C1:C2\')',
      resultCheck: (payload: any) => {
        assert.deepEqual(payload.formulas, [['=A1+B1'], ['=A2+B2']]);
      },
    },
    {
      name: 'set_formulas',
      args: { itemId: 'excel-1', operation: 'set_formulas', worksheet: 'Sheet1', range: 'C1:C2', formulas: [['=A1+B1'], ['=A2+B2']] },
      method: 'patch',
      endpoint: '/me/drive/items/excel-1/workbook/worksheets/Sheet1/range(address=\'C1:C2\')',
      resultCheck: (payload: any) => {
        assert.equal(payload.formulasSet, 2);
        assert.equal(payload.address, 'Sheet1!C1:C2');
      },
    },
    {
      name: 'create_table',
      args: { itemId: 'excel-1', operation: 'create_table', worksheet: 'Sheet1', range: 'A1:B4', tableName: 'Sales', hasHeaders: false },
      method: 'post',
      endpoint: '/me/drive/items/excel-1/workbook/worksheets/Sheet1/tables/add',
      resultCheck: (payload: any, mock: ReturnType<typeof createMockGraphClient>) => {
        assert.equal(payload.table.name, 'Sales');
        assert.equal(mock.methodCalls('patch').length, 1);
        assert.deepEqual(mock.methodCalls('patch')[0].args, [
          '/me/drive/items/excel-1/workbook/tables/table-1',
          { name: 'Sales' },
          { headers: undefined },
        ]);
      },
    },
    {
      name: 'create_chart',
      args: { itemId: 'excel-1', operation: 'create_chart', worksheet: 'Sheet1', range: 'A1:B4', chartType: 'LineMarkers' },
      method: 'post',
      endpoint: '/me/drive/items/excel-1/workbook/worksheets/Sheet1/charts/add',
      resultCheck: (payload: any) => {
        assert.equal(payload.chart.type, 'LineMarkers');
        assert.equal(payload.chart.id, 'chart-1');
      },
    },
  ] as const;

  for (const scenario of scenarios) {
    const mock = createMockGraphClient({
      get: async (endpoint: string, params?: Record<string, unknown>) => {
        if (endpoint === '/me/drive/items/excel-1/workbook/worksheets') {
          return {
            success: true,
            data: {
              value: [
                { id: 'ws-1', name: 'Sheet1', position: 0, visibility: 'Visible' },
                { id: 'ws-2', name: 'Backlog', position: 1, visibility: 'Visible' },
              ],
            },
          };
        }

        if (endpoint === '/me/drive/items/excel-1/workbook/worksheets/Sheet1/range(address=\'C1:C2\')') {
          assert.deepEqual(params, { '$select': 'formulas,address,rowCount,columnCount' });
          return {
            success: true,
            data: {
              address: 'Sheet1!C1:C2',
              rowCount: 2,
              columnCount: 1,
              formulas: [['=A1+B1'], ['=A2+B2']],
            },
          };
        }

        throw new Error(`Unexpected GET ${endpoint}`);
      },
      post: async (endpoint: string, data?: any, options?: Record<string, unknown>) => {
        assert.deepEqual(options, { headers: undefined });

        if (endpoint === '/me/drive/items/excel-1/workbook/worksheets/add') {
          assert.deepEqual(data, { name: 'Backlog' });
          return { success: true, data: { id: 'ws-2', name: 'Backlog', position: 1, visibility: 'Visible' } };
        }

        if (endpoint === '/me/drive/items/excel-1/workbook/worksheets/Sheet1/tables/add') {
          assert.deepEqual(data, { address: 'Sheet1!A1:B4', hasHeaders: false });
          return {
            success: true,
            data: { id: 'table-1', name: 'Table1', showHeaders: true, showTotals: false, style: 'TableStyleMedium2' },
          };
        }

        if (endpoint === '/me/drive/items/excel-1/workbook/worksheets/Sheet1/charts/add') {
          assert.deepEqual(data, { type: 'LineMarkers', sourceData: 'Sheet1!A1:B4', seriesBy: 'auto' });
          return {
            success: true,
            data: { id: 'chart-1', name: 'SalesChart', height: 320, width: 480, top: 5, left: 10 },
          };
        }

        throw new Error(`Unexpected POST ${endpoint}`);
      },
      patch: async (endpoint: string, data?: any, options?: Record<string, unknown>) => {
        assert.deepEqual(options, { headers: undefined });

        if (endpoint === '/me/drive/items/excel-1/workbook/worksheets/Sheet1/range(address=\'A1:B2\')') {
          assert.deepEqual(data, { values: [[1, 2], [3, 4]] });
          return { success: true, data: { address: 'Sheet1!A1:B2' } };
        }

        if (endpoint === '/me/drive/items/excel-1/workbook/worksheets/Sheet1/range(address=\'C1:C2\')') {
          assert.deepEqual(data, { formulas: [['=A1+B1'], ['=A2+B2']] });
          return { success: true, data: { address: 'Sheet1!C1:C2' } };
        }

        if (endpoint === '/me/drive/items/excel-1/workbook/tables/table-1') {
          assert.deepEqual(data, { name: 'Sales' });
          return { success: true, data: { id: 'table-1' } };
        }

        throw new Error(`Unexpected PATCH ${endpoint}`);
      },
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleExcelOperations(scenario.args)) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(response.isError, undefined, scenario.name);
    assert.equal(payload.operation, scenario.name);
    scenario.resultCheck(payload, mock);

    const calls = mock.calls.filter((call) => call.method === scenario.method);
    assert.ok(calls.some((call) => call.args[0] === scenario.endpoint), scenario.name);
  }
});

test('excel_operations validates required fields per operation and rejects invalid operations', async () => {
  const validationCases = [
    {
      args: { itemId: 'excel-1', operation: 'read_range' },
      error: 'Range is required for read_range operation',
    },
    {
      args: { itemId: 'excel-1', operation: 'write_range', range: 'A1:B1' },
      error: 'Range and values are required for write_range operation',
    },
    {
      args: { itemId: 'excel-1', operation: 'get_formulas' },
      error: 'Range is required for get_formulas operation',
    },
    {
      args: { itemId: 'excel-1', operation: 'set_formulas', range: 'A1' },
      error: 'Range and formulas are required for set_formulas operation',
    },
    {
      args: { itemId: 'excel-1', operation: 'create_table', range: 'A1:B2' },
      error: 'Range and tableName are required for create_table operation',
    },
    {
      args: { itemId: 'excel-1', operation: 'create_chart', range: 'A1:B2' },
      error: 'Range and chartType are required for create_chart operation',
    },
    {
      args: { itemId: 'excel-1', operation: 'invalid_operation' },
      error: 'Invalid operation: invalid_operation',
    },
    {
      args: { operation: 'list_worksheets' },
      error: 'Either itemId or itemPath is required',
    },
  ] as const;

  for (const validationCase of validationCases) {
    const mock = createMockGraphClient();
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleExcelOperations(validationCase.args)) as ToolEnvelope;
    assert.equal(response.isError, true, validationCase.error);
    assert.match(response.content[0].text, new RegExp(validationCase.error.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')));
  }
});

test('excel_analysis covers statistics, pivot summary, data validation, named ranges, and used range', async () => {
  const scenarios = [
    {
      name: 'statistics',
      args: { itemId: 'excel-1', analysisType: 'statistics', worksheet: 'Sheet1', range: 'A1:B3' },
      get: async (endpoint: string) => {
        assert.equal(endpoint, '/me/drive/items/excel-1/workbook/worksheets/Sheet1/range(address=\'A1:B3\')');
        return {
          success: true,
          data: {
            address: 'Sheet1!A1:B3',
            rowCount: 3,
            columnCount: 2,
            values: [[1, 2], [3, 4], [5, 'x']],
          },
        };
      },
      check: (payload: any) => {
        assert.equal(payload.statistics.length, 2);
        assert.equal(payload.statistics[0].sum, 9);
        assert.equal(payload.statistics[1].max, 4);
      },
    },
    {
      name: 'pivot_summary',
      args: { itemId: 'excel-1', analysisType: 'pivot_summary', worksheet: 'Pivot' },
      get: async (endpoint: string) => {
        assert.equal(endpoint, '/me/drive/items/excel-1/workbook/worksheets/Pivot/pivotTables');
        return {
          success: true,
          data: { value: [{ id: 'pivot-1', name: 'PivotSales', worksheet: { name: 'Pivot' } }] },
        };
      },
      post: async (endpoint: string) => {
        assert.equal(endpoint, '/me/drive/items/excel-1/workbook/worksheets/Pivot/pivotTables/pivot-1/refresh');
        return { success: true };
      },
      check: (payload: any, mock: ReturnType<typeof createMockGraphClient>) => {
        assert.equal(payload.pivotTableCount, 1);
        assert.equal(payload.pivotTables[0].id, 'pivot-1');
        assert.equal(mock.methodCalls('post').length, 1);
      },
    },
    {
      name: 'data_validation',
      args: { itemId: 'excel-1', analysisType: 'data_validation', worksheet: 'Sheet1' },
      get: async (endpoint: string) => {
        assert.equal(endpoint, '/me/drive/items/excel-1/workbook/worksheets/Sheet1/usedRange/dataValidation');
        return {
          success: true,
          data: {
            errorAlert: 'stop',
            errorMessage: 'Choose a listed value',
            errorTitle: 'Invalid',
            operator: 'between',
            type: 'list',
            formula1: '"A,B"',
            formula2: null,
          },
        };
      },
      check: (payload: any) => {
        assert.equal(payload.validation.type, 'list');
        assert.equal(payload.range, 'usedRange');
      },
    },
    {
      name: 'named_ranges',
      args: { itemPath: 'Reports/book.xlsx', siteId: 'site-1', analysisType: 'named_ranges' },
      get: async (endpoint: string) => {
        if (endpoint === '/sites/site-1/drive/root:/Reports/book.xlsx') {
          return { success: true, data: { id: 'excel-1' } };
        }

        assert.equal(endpoint, '/sites/site-1/drive/items/excel-1/workbook/names');
        return {
          success: true,
          data: {
            value: [
              { name: 'InputRange', type: 'Range', value: 'Sheet1!A1:B5', visible: true, scope: 'Workbook' },
            ],
          },
        };
      },
      check: (payload: any) => {
        assert.equal(payload.count, 1);
        assert.equal(payload.namedRanges[0].name, 'InputRange');
      },
    },
    {
      name: 'used_range',
      args: { itemId: 'excel-1', analysisType: 'used_range', worksheet: 'Sheet1' },
      get: async (endpoint: string) => {
        assert.equal(endpoint, '/me/drive/items/excel-1/workbook/worksheets/Sheet1/usedRange');
        return {
          success: true,
          data: {
            address: 'Sheet1!A1:B2',
            rowCount: 2,
            columnCount: 2,
            values: [[1, 'x'], ['', null]],
            formulas: [['=A1*2', null], [null, null]],
          },
        };
      },
      check: (payload: any) => {
        assert.equal(payload.usedRange.totalCells, 4);
        assert.deepEqual(payload.usedRange.cellTypes, {
          empty: 2,
          numbers: 1,
          text: 1,
          formulas: 1,
        });
      },
    },
  ] as const;

  for (const scenario of scenarios) {
    const mock = createMockGraphClient({
      get: scenario.get,
      post: (scenario as any).post,
    });
    __setGraphClientInstanceForTests(mock.client as any);

    const response = (await handleExcelAnalysis(scenario.args)) as ToolEnvelope;
    const payload = parsePayload<any>(response);

    assert.equal(response.isError, undefined, scenario.name);
    assert.equal(payload.analysisType, scenario.name);
    scenario.check(payload, mock);
  }
});

test('excel_analysis rejects missing item reference and invalid analysis type', async () => {
  let mock = createMockGraphClient();
  __setGraphClientInstanceForTests(mock.client as any);

  const missingItemResponse = (await handleExcelAnalysis({ analysisType: 'used_range' })) as ToolEnvelope;
  assert.equal(missingItemResponse.isError, true);
  assert.match(missingItemResponse.content[0].text, /Either itemId or itemPath is required/);

  mock = createMockGraphClient();
  __setGraphClientInstanceForTests(mock.client as any);

  const invalidTypeResponse = (await handleExcelAnalysis({ itemId: 'excel-1', analysisType: 'bogus' })) as ToolEnvelope;
  assert.equal(invalidTypeResponse.isError, true);
  assert.match(invalidTypeResponse.content[0].text, /Invalid analysis type: bogus/);
});

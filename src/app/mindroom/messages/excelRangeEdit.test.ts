import { describe, expect, it } from 'vitest';
import { EXCEL_RANGE_EDIT_TOOL, excelDesktopUrl, parseExcelRangeEdit } from './excelRangeEdit';

const WEB_URL = 'https://contoso.sharepoint.com/sites/finance/Shared Documents/Forecast.xlsx';

const edit = (overrides: Record<string, unknown> = {}) => ({
  file: '/drives/b!1/items/01X',
  worksheet: 'Budget',
  address: 'B2:C3',
  values: [
    ['Revenue', 1200],
    [null, '=SUM(C1:C2)'],
  ],
  expected_formulas: [
    ['Revenue', 1000],
    ['', '=SUM(C1:C2)'],
  ],
  expected_web_url: WEB_URL,
  ...overrides,
});

describe('parseExcelRangeEdit', () => {
  it('lists only the cells an edit changes, with their addresses', () => {
    const review = parseExcelRangeEdit(EXCEL_RANGE_EDIT_TOOL, edit());

    expect(review).toEqual({
      worksheet: 'Budget',
      address: 'B2:C3',
      webUrl: WEB_URL,
      changes: [{ cell: 'C2', before: 1000, after: 1200, numberFormat: null }],
      unchangedCount: 3,
    });
  });

  it('starts at the top-left cell of a reversed or absolute address', () => {
    const review = parseExcelRangeEdit(
      EXCEL_RANGE_EDIT_TOOL,
      edit({
        address: '$C$3:B2',
        values: [
          [1, null],
          [null, null],
        ],
      })
    );

    expect(review?.changes.map((change) => change.cell)).toEqual(['B2']);
  });

  it('names cells past column Z', () => {
    const review = parseExcelRangeEdit(
      EXCEL_RANGE_EDIT_TOOL,
      edit({ address: 'AA10', values: [[2]], expected_formulas: [[1]] })
    );

    expect(review?.changes[0].cell).toBe('AA10');
  });

  it('shows a number format change on a cell whose value stays', () => {
    const review = parseExcelRangeEdit(
      EXCEL_RANGE_EDIT_TOOL,
      edit({
        number_formats: [
          [null, '0.0%'],
          [null, null],
        ],
        values: [
          [null, null],
          [null, null],
        ],
      })
    );

    expect(review?.changes).toEqual([
      { cell: 'C2', before: 1000, after: null, numberFormat: '0.0%' },
    ]);
  });

  it('never shows a cell as unchanged when its type or redacted text differs or may differ', () => {
    const review = parseExcelRangeEdit(
      EXCEL_RANGE_EDIT_TOOL,
      edit({
        address: 'A1:C1',
        values: [['42', 'https://x.example/?token=***redacted***', '=sum(a1)']],
        expected_formulas: [[42, 'https://x.example/?token=***redacted***', '=SUM(A1)']],
      })
    );

    expect(review?.changes.map((change) => change.cell)).toEqual(['A1', 'B1', 'C1']);
    expect(review?.unchangedCount).toBe(0);
  });

  it.each([
    ['another tool', 'microsoft_files_upload_file', {}],
    ['a missing worksheet', EXCEL_RANGE_EDIT_TOOL, { worksheet: undefined }],
    ['an address with a worksheet', EXCEL_RANGE_EDIT_TOOL, { address: 'Budget!B2:C3' }],
    ['values of another shape', EXCEL_RANGE_EDIT_TOOL, { values: [['x', 1]] }],
    ['ragged values', EXCEL_RANGE_EDIT_TOOL, { values: [['x', 1], ['y']] }],
    [
      'before of another shape',
      EXCEL_RANGE_EDIT_TOOL,
      {
        expected_formulas: [
          [1, 2, 3],
          [4, 5, 6],
        ],
      },
    ],
    [
      'a null before',
      EXCEL_RANGE_EDIT_TOOL,
      {
        expected_formulas: [
          [null, 1],
          ['', 2],
        ],
      },
    ],
    [
      'a nested cell',
      EXCEL_RANGE_EDIT_TOOL,
      {
        values: [
          [{ x: 1 }, 1],
          [null, 2],
        ],
      },
    ],
    ['formats of another shape', EXCEL_RANGE_EDIT_TOOL, { number_formats: [['@']] }],
  ])('leaves %s to the generic argument view', (_name, toolName, overrides) => {
    expect(parseExcelRangeEdit(toolName, edit(overrides))).toBeNull();
  });

  it('leaves missing arguments to the generic argument view', () => {
    expect(parseExcelRangeEdit(EXCEL_RANGE_EDIT_TOOL, undefined)).toBeNull();
  });

  it.each([
    'http://contoso.sharepoint.com/Forecast.xlsx',
    'https://contoso.sharepoint.com.evil.example/Forecast.xlsx',
    'https://user:pass@contoso.sharepoint.com/Forecast.xlsx',
    'not a url',
  ])('links no workbook outside OneDrive and SharePoint: %s', (url) => {
    expect(
      parseExcelRangeEdit(EXCEL_RANGE_EDIT_TOOL, edit({ expected_web_url: url }))?.webUrl
    ).toBeNull();
  });

  it.each([
    'https://onedrive.live.com/edit?id=1',
    'https://1drv.ms/x/s!abc',
    'https://contoso-my.sharepoint.com/personal/a/x.xlsx',
  ])('links a workbook in OneDrive or SharePoint: %s', (url) => {
    expect(
      parseExcelRangeEdit(EXCEL_RANGE_EDIT_TOOL, edit({ expected_web_url: url }))?.webUrl
    ).toBe(url);
  });
});

describe('excelDesktopUrl', () => {
  it('opens the workbook for editing in desktop Excel', () => {
    expect(excelDesktopUrl(WEB_URL)).toBe(`ms-excel:ofe|u|${WEB_URL}`);
  });
});

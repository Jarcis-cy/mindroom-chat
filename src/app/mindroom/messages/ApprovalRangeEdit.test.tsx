import React from 'react';
import { createInstance } from 'i18next';
import { I18nextProvider } from 'react-i18next';
import { act, create, ReactTestRenderer } from 'react-test-renderer';
import { afterEach, describe, expect, it, vi } from 'vitest';
import en from '../../locales/en.json';
import { ApprovalArguments } from './ApprovalArguments';
import { parseToolApprovalContent } from './toolApproval';

vi.mock('../../hooks/useMatrixClient', () => ({ useMatrixClient: () => ({}) }));
vi.mock('./sidecarDownload', () => ({ downloadMindroomSidecarBlob: vi.fn() }));
vi.mock('./ApprovalArgumentFields.css', () => ({
  Fields: 'Fields',
  Key: 'Key',
  Value: 'Value',
  Clamped: 'Clamped',
  Scalar: 'Scalar',
  Code: 'Code',
  Toggle: 'Toggle',
  Count: 'Count',
  Arguments: 'Arguments',
}));
vi.mock('./ApprovalRangeEdit.css', () => ({
  RangeEdit: 'RangeEdit',
  Header: 'Header',
  Target: 'Target',
  Links: 'Links',
  Scroll: 'Scroll',
  Table: 'Table',
  Cell: 'Cell',
  Value: 'Value',
  Before: 'Before',
  Muted: 'Muted',
  Format: 'Format',
}));

const WEB_URL = 'https://contoso.sharepoint.com/sites/finance/Forecast.xlsx';
const ARGUMENTS = {
  file: '/drives/b!1/items/01X',
  worksheet: 'Q3 Budget',
  address: 'B2:C2',
  values: [['0042', true]],
  expected_formulas: [['', false]],
  expected_web_url: WEB_URL,
  number_formats: [['@', null]],
};

const approvalFor = (content: Record<string, unknown>) =>
  parseToolApprovalContent('io.mindroom.tool_approval', {
    approval_id: 'one',
    tool_name: 'microsoft_files_excel_update_range',
    agent_name: 'assistant',
    status: 'pending',
    requested_at: '2026-11-02T12:00:00Z',
    expires_at: '2026-11-02T12:30:00Z',
    ...content,
  })!;

let renderer: ReactTestRenderer;
afterEach(() => act(() => renderer?.unmount()));

const render = (approval: ReturnType<typeof approvalFor>) => {
  const i18n = createInstance();
  void i18n.init({ lng: 'en', resources: { en: { translation: en } }, initAsync: false });
  act(() => {
    renderer = create(
      <I18nextProvider i18n={i18n}>
        <ApprovalArguments approval={approval} />
      </I18nextProvider>
    );
  });
  return JSON.stringify(renderer.toJSON());
};

describe('Excel range edit approvals', () => {
  it('show each changed cell before and after, with links to the workbook', () => {
    const text = render(approvalFor({ arguments: ARGUMENTS }));

    expect(text).toContain("'Q3 Budget'!B2:C2");
    for (const part of ['B2', 'empty', '0042', 'Format: @', 'C2', 'FALSE', 'TRUE'])
      expect(text).toContain(part);
    const hrefs = renderer.root.findAllByType('a').map((link) => link.props.href);
    expect(hrefs).toEqual([`ms-excel:ofe|u|${WEB_URL}`, WEB_URL]);
  });

  it('wait for the complete arguments when the card shows a truncated preview', () => {
    const text = render(
      approvalFor({
        arguments: { address: 'B2:C2' },
        arguments_truncated: true,
        full_arguments_url: 'mxc://x/y',
      })
    );

    expect(text).not.toContain('Before');
  });

  it('use complete arguments the card carries next to a truncated preview', () => {
    const text = render(
      approvalFor({
        arguments: { address: 'B2:C2' },
        arguments_truncated: true,
        full_arguments: ARGUMENTS,
      })
    );

    expect(text).toContain('Before');
    expect(text).toContain('0042');
  });

  it('leave other tools to the argument fields', () => {
    const text = render(
      approvalFor({ tool_name: 'microsoft_files_upload_file', arguments: ARGUMENTS })
    );

    expect(text).not.toContain('Before');
  });
});

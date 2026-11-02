import { useTranslation } from 'react-i18next';
import React from 'react';
import { Text } from 'folds';
import classNames from 'classnames';
import { ExcelCellValue, ExcelRangeEditReview, excelDesktopUrl } from './excelRangeEdit';
import * as css from './ApprovalRangeEdit.css';

// Enough rows to review by eye; the full argument view lists every cell.
const MAX_SHOWN_CHANGES = 200;

// Excel's own notation: a sheet name with anything but letters, digits, and underscores is quoted.
const rangeLabel = (worksheet: string, address: string) =>
  /^[\p{L}\p{N}_]+$/u.test(worksheet)
    ? `${worksheet}!${address}`
    : `'${worksheet.replace(/'/g, "''")}'!${address}`;

function CellValue({ value }: { value: ExcelCellValue }) {
  const { t } = useTranslation();
  if (value === '')
    return <span className={css.Muted}>{t('mindroomUi.messages.approvalRangeEdit.empty')}</span>;
  if (typeof value === 'boolean') return <>{value ? 'TRUE' : 'FALSE'}</>;
  return <>{String(value)}</>;
}

/** The cells an Excel range edit changes, before and after, with links to the workbook. */
export function ApprovalRangeEdit({ review }: { review: ExcelRangeEditReview }) {
  const { t } = useTranslation();
  const shown = review.changes.slice(0, MAX_SHOWN_CHANGES);
  return (
    <div className={css.RangeEdit}>
      <div className={css.Header}>
        <Text size="T300" className={css.Target}>
          {rangeLabel(review.worksheet, review.address)}
        </Text>
        {review.webUrl && (
          <span className={css.Links}>
            <a href={excelDesktopUrl(review.webUrl)}>
              {t('mindroomUi.messages.approvalRangeEdit.openInExcel')}
            </a>
            <a href={review.webUrl} target="_blank" rel="noreferrer noopener">
              {t('mindroomUi.messages.approvalRangeEdit.openInBrowser')}
            </a>
          </span>
        )}
      </div>
      {shown.length > 0 ? (
        <div className={css.Scroll}>
          <table className={css.Table}>
            <thead>
              <tr>
                <th scope="col">{t('mindroomUi.messages.approvalRangeEdit.cell')}</th>
                <th scope="col">{t('mindroomUi.messages.approvalRangeEdit.before')}</th>
                <th scope="col">{t('mindroomUi.messages.approvalRangeEdit.after')}</th>
              </tr>
            </thead>
            <tbody>
              {shown.map((change) => (
                <tr key={change.cell}>
                  <th scope="row" className={css.Cell}>
                    {change.cell}
                  </th>
                  <td className={classNames(css.Value, css.Before)}>
                    <CellValue value={change.before} />
                  </td>
                  <td className={css.Value}>
                    {change.after === null ? (
                      <span className={css.Muted}>
                        {t('mindroomUi.messages.approvalRangeEdit.valueKept')}
                      </span>
                    ) : (
                      <CellValue value={change.after} />
                    )}
                    {change.numberFormat !== null && (
                      <span className={css.Format}>
                        {t('mindroomUi.messages.approvalRangeEdit.numberFormat', {
                          format: change.numberFormat,
                        })}
                      </span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <Text size="T200">{t('mindroomUi.messages.approvalRangeEdit.noChanges')}</Text>
      )}
      {review.changes.length > shown.length && (
        <Text size="T200">
          {t('mindroomUi.messages.approvalRangeEdit.moreChanges', {
            count: review.changes.length - shown.length,
          })}
        </Text>
      )}
      {review.unchangedCount > 0 && (
        <Text size="T200" className={css.Muted}>
          {t('mindroomUi.messages.approvalRangeEdit.unchangedCells', {
            count: review.unchangedCount,
          })}
        </Text>
      )}
    </div>
  );
}

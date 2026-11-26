import { expect, test } from '@playwright/test';
import { parseHwpCenteredCellLabels } from '../src/hwp-centered-cell-labels';
import { hwpWrapFingerprint } from '../src/hwp-text-advances';

test('a first-fragment label remains distinct from the stored two-label policy', () => {
  const fingerprint = hwpWrapFingerprint(JSON.stringify(['07 안내 항목']));
  const leading = '67.3 ' + fingerprint;
  expect(parseHwpCenteredCellLabels(leading)).toEqual({ width: 67.3, fingerprint });
  expect(parseHwpCenteredCellLabels(leading + ' first')).toEqual({ width: 67.3, fingerprint, firstFragment: true });
  expect(parseHwpCenteredCellLabels(leading + ' last')).toBeUndefined();
  expect(parseHwpCenteredCellLabels('0 ' + fingerprint + ' first')).toBeUndefined();
});

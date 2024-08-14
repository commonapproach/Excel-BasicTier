import { Indicator } from './Indicator';
import { IndicatorReport } from './IndicatorReport';
import { Organization } from './Organization';
import { Outcome } from './Outcome';
import { Theme } from './Theme';

export const map = {
  Organization,
  Theme,
  Outcome,
  Indicator,
  IndicatorReport,
};

export type ModelType = keyof typeof map;

export function createInstance<T extends ModelType>(sheetName: T) {
  const Model = map[sheetName];
  if (!Model) {
    throw new Error(`Model for ${sheetName} not found.`);
  }
  return new Model();
}

export const ignoredFields: { [key: string]: string[] } = {
  Theme: ['hasOutcome'],
  Outcome: ['forOrganization'],
  Indicator: ['forOrganization', 'forOutcome'],
};

export * from './Indicator';
export * from './IndicatorReport';
export * from './Organization';
export * from './Outcome';
export * from './Theme';

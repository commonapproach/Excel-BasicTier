import { CellDataInterface } from './cell.interface';

export type RecordInterface = {
  fields: {
    '@id': {
      type: string;
      data: string;
      createdAt: string;
      isPrimary: boolean;
    };
    [key: string]: {
      '@context'?: string;
      type: string;
      data: CellDataInterface;
      createdAt: string;
      isPrimary: boolean;
    };
  };
  recordId: string;
};

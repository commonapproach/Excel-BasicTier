import { Base } from './Base';
import { IndicatorReport } from './IndicatorReport';

export class Indicator extends Base {
  static className = 'Indicator';
  protected _fields = [
    {
      name: '@id',
      type: 'string',
      representedType: 'string',
      primary: true,
      unique: true,
      notNull: true,
      required: true,
      semiRequired: false,
    },
    {
      name: 'hasName',
      type: 'string',
      representedType: 'string',
      unique: true,
      notNull: true,
      required: true,
      semiRequired: false,
    },
    {
      name: 'hasDescription',
      type: 'text',
      representedType: 'string',
      defaultValue: '',
      unique: false,
      notNull: false,
      required: false,
      semiRequired: false,
    },
    {
      name: 'hasIndicatorReport',
      type: 'link',
      representedType: 'array',
      defaultValue: [],
      link: IndicatorReport,
      unique: false,
      notNull: false,
      required: false,
      semiRequired: false,
    },
  ];
}

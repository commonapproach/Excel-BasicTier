import { Base } from './Base';
import { Indicator } from './Indicator';

export class IndicatorReport extends Base {
  static className = 'IndicatorReport';
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
      name: 'hasComment',
      type: 'string',
      representedType: 'string',
      defaultValue: '',
      unique: false,
      notNull: false,
      required: false,
      semiRequired: false,
    },
    {
      name: 'i72:value',
      type: 'i72',
      representedType: 'string',
      defaultValue: '',
      unique: false,
      notNull: true,
      required: true,
      semiRequired: false,
    },
    {
      name: 'i72:unit_of_measure',
      type: 'i72',
      representedType: 'string',
      defaultValue: '',
      unique: false,
      notNull: false,
      required: false,
      semiRequired: false,
    },
    {
      name: 'forIndicator',
      type: 'link',
      representedType: 'string',
      defaultValue: '',
      link: Indicator,
      unique: false,
      notNull: false,
      required: false,
      semiRequired: true,
    },
  ];
}

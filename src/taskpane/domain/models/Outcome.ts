import { Base } from './Base';
import { Indicator } from './Indicator';
import { Theme } from './Theme';

export class Outcome extends Base {
  static className = 'Outcome';
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
      notNull: true,
      required: true,
      semiRequired: false,
    },
    {
      name: 'hasIndicator',
      type: 'link',
      representedType: 'array',
      defaultValue: [],
      link: Indicator,
      unique: false,
      notNull: false,
      required: false,
      semiRequired: true,
    },
    {
      name: 'forTheme',
      type: 'link',
      representedType: 'array',
      defaultValue: [],
      link: Theme,
      unique: false,
      notNull: false,
      required: false,
      semiRequired: true,
    },
  ];
}

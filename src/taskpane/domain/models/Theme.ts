import { Base } from './Base';

export class Theme extends Base {
  static className = 'Theme';
  protected _fields = [
    {
      name: '@id',
      type: 'string',
      representedType: 'string',
      primary: true,
      unique: false,
      notNull: true,
      required: false,
      semiRequired: true,
    },
    {
      name: 'hasName',
      type: 'string',
      representedType: 'string',
      unique: false,
      notNull: true,
      required: false,
      semiRequired: true,
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
  ];
}

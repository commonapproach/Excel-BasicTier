import { Base } from './Base';
import { Indicator } from './Indicator';
import { Outcome } from './Outcome';

export class Organization extends Base {
  static className = 'Organization';
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
      name: 'org:hasLegalName',
      type: 'string',
      representedType: 'string',
      unique: true,
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
      name: 'hasOutcome',
      type: 'link',
      representedType: 'array',
      defaultValue: [],
      link: Outcome,
      unique: false,
      notNull: false,
      required: false,
      semiRequired: true,
    },
  ];
}

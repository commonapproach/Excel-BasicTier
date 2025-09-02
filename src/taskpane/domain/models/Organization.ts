import { Address } from "./Address";
import { Base } from "./Base";
import { Indicator } from "./Indicator";
import { Outcome } from "./Outcome";

export class Organization extends Base {
  static className = "Organization";

  constructor() {
    super();
    this._fields = [
      {
        name: "@id",
        type: "string",
        representedType: "string",
        primary: true,
        unique: true,
        notNull: true,
        required: true,
        semiRequired: false,
      },
      {
        name: "org:hasLegalName",
        displayName: "hasLegalName",
        type: "string",
        representedType: "string",
        unique: true,
        notNull: true,
        required: true,
        semiRequired: false,
      },
      {
        name: "hasAddress",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Address, field: "forOrganization" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasIndicator",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Indicator, field: "forOrganization" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "hasOutcome",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Outcome, field: "forOrganization" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
    ];
  }
}

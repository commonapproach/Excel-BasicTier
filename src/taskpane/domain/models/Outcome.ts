import { Base } from "./Base";
import { Indicator } from "./Indicator";
import { Organization } from "./Organization";
import { Theme } from "./Theme";

export class Outcome extends Base {
  static className = "Outcome";

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
        name: "org:hasName",
        displayName: "hasName",
        type: "string",
        representedType: "string",
        unique: true,
        notNull: true,
        required: true,
        semiRequired: false,
      },
      {
        name: "hasDescription",
        type: "text",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "forTheme",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Theme, field: "hasOutcome" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "forOrganization",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Organization, field: "hasOutcome" },
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
        link: { table: Indicator, field: "forOutcome" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
    ];
  }
}

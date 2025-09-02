import { getUnitOptions } from "../fetchServer/getUnitsOfMeasure";
import { Base } from "./Base";
import { IndicatorReport } from "./IndicatorReport";
import { Organization } from "./Organization";
import { Outcome } from "./Outcome";
import { Population } from "./Population";
import { Theme } from "./Theme";

export class Indicator extends Base {
  static className: string = "Indicator";

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
        name: "unitDescription",
        displayName: "unitDescription",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "i72:unit_of_measure",
        displayName: "unit_of_measure",
        type: "select",
        representedType: "string",
        defaultValue: "",
        selectOptions: [],
        getOptionsAsync: async () => {
          const opts = await getUnitOptions();
          return opts.map((o) => ({ id: o.id, name: o.name }));
        },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "describesPopulation",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Population, field: "forIndicator" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "forOrganization",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Organization, field: "hasIndicator" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "forOutcome",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Outcome, field: "hasIndicator" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "forTheme",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Theme, field: "hasIndicator" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasIndicatorReport",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: IndicatorReport, field: "forIndicator" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
    ];
  }
}

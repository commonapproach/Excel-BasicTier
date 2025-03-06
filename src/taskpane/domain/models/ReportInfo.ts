import { Base } from "./Base";
import { Organization } from "./Organization";

export class ReportInfo extends Base {
  static className: string = "ReportInfo";

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
        name: "hasName",
        type: "string",
        representedType: "string",
        unique: true,
        notNull: true,
        required: true,
        semiRequired: false,
      },
      {
        name: "forOrganization",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Organization, field: "hasReportInfo" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "prov:startedAtTime",
        displayName: "startedAtTime",
        type: "datetime",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "prov:endedAtTime",
        displayName: "endedAtTime",
        type: "datetime",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
    ];
  }
}

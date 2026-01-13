import { Base } from "./Base";
import { FundingState } from "./FundingState";
import { Organization } from "./Organization";

export class FundingStatus extends Base {
  static className = "FundingStatus";

  constructor() {
    super();
    this._fields = [
      {
        name: "@id",
        type: "string",
        representedType: "string",
        primary: true,
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "forOrganization",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Organization, field: "hasFundingStatus" },
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "forFunder",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "hasFundingState",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: FundingState, field: "forFundingStatus" },
        unique: false,
        notNull: false,
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
        name: "reportedDate",
        type: "datetime",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
    ];
  }
}

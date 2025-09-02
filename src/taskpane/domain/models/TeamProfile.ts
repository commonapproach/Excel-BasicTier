import { Base } from "./Base";
import { EDGProfile } from "./EDGProfile";

export class TeamProfile extends Base {
  static className = "TeamProfile";

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
        name: "hasTeamSize",
        type: "number",
        representedType: "number",
        defaultValue: 0,
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasEDGSize",
        type: "number",
        representedType: "number",
        defaultValue: 0,
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "hasEDGProfile",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: EDGProfile, field: "forTeamProfile" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "hasComment",
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
        type: "date",
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

import { Base } from "./Base";
import { Characteristic } from "./Characteristic";

export class EDGProfile extends Base {
  static className = "EDGProfile";

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
        name: "forEDG",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Characteristic, field: "hasEDGProfile" },
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasSize",
        type: "number",
        representedType: "number",
        defaultValue: 0,
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "reportedDate",
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

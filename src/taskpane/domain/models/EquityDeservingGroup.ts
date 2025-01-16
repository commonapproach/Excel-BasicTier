import { Base } from "./Base";
import { Characteristic } from "./Characteristic";

export class EquityDeservingGroup extends Base {
  static className = "EquityDeservingGroup";

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
        name: "hasCharacteristic",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Characteristic, field: "forEquityDeservingGroup" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "isDefined",
        type: "boolean",
        representedType: "boolean",
        defaultValue: false,
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
    ];
  }
}

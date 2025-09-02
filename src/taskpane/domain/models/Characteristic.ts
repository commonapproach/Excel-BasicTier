import { Base } from "./Base";
import { PopulationServed } from "./PopulationServed";

export class Characteristic extends Base {
  static className = "Characteristic";

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
        name: "org:hasName",
        displayName: "hasName",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasValue",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasCode",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: PopulationServed, field: "forCharacteristic" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
    ];
  }
}

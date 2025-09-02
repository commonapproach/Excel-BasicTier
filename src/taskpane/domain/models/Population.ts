import { Base } from "./Base";

export class Population extends Base {
  static className: string = "Population";

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
        name: "rdfs:label",
        displayName: "label",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: true,
        required: true,
        semiRequired: false,
      },
    ];
  }
}

export default Population;

import { Base } from "./Base";

export class PopulationServed extends Base {
  static className = "PopulationServed";

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
        name: "@type",
        type: "string",
        representedType: "array",
        defaultValue: [],
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "hasIdentifier",
        type: "string",
        representedType: "string",
        defaultValue: "",
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
        name: "hasDescription",
        type: "text",
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

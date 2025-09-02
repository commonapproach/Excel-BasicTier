import { Base } from "./Base";

export class Theme extends Base {
  static className = "Theme";

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
        name: "hasCode",
        type: "string",
        representedType: "array",
        defaultValue: [],
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "relatesTo",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Theme, field: "relatesTo" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
    ];
  }
}

import { Base } from "./Base";

export class Person extends Base {
  static className = "Person";

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
        name: "foaf:givenName",
        displayName: "givenName",
        type: "string",
        representedType: "string",
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "foaf:familyName",
        displayName: "familyName",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: true,
        required: false,
        semiRequired: true,
      },
      {
        name: "ic:hasEmail",
        displayName: "hasEmail",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
    ];
  }
}

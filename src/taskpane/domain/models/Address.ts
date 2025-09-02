import { Base } from "./Base";

export class Address extends Base {
  static className: string = "Address";

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
        name: "streetAddress",
        displayName: "streetAddress",
        type: "text",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: true,
        semiRequired: false,
      },
      {
        name: "extendedAddress",
        displayName: "extendedAddress",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "addressLocality",
        displayName: "addressLocality",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "addressRegion",
        displayName: "addressRegion",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "postalCode",
        displayName: "postalCode",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "addressCountry",
        displayName: "addressCountry",
        type: "string",
        representedType: "string",
        defaultValue: "",
        unique: false,
        notNull: false,
        required: false,
        semiRequired: false,
      },
      {
        name: "postOfficeBoxNumber",
        displayName: "postOfficeBoxNumber",
        type: "string",
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

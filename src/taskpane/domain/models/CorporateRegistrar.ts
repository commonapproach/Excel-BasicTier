import { Base } from "./Base";

export class CorporateRegistrar extends Base {
	static className: string = "CorporateRegistrar";

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
				name: "hasIdentifier",
				type: "string",
				representedType: "string",
				defaultValue: "",
				unique: false,
				notNull: false,
				required: false,
				semiRequired: false,
			},
			{
				name: "cids:hasName",
				displayName: "hasName",
				type: "string",
				representedType: "string",
				unique: false,
				notNull: true,
				required: true,
				semiRequired: false,
			},
			{
				name: "cids:hasDescription",
				displayName: "hasDescription",
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
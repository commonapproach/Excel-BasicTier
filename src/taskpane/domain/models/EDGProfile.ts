
import { Base } from "./Base";
import { EquityDeservingGroup } from "./EquityDeservingGroup";

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
				link: { table: EquityDeservingGroup, field: "hasEDGProfile" },
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

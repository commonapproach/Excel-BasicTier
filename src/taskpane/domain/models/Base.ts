export class Base {
  protected _fields: FieldType[] = [];
  public getFieldByName(name: string): FieldType {
    const field = this.getAllFields().find((f) => {
      if (f.name === name) return true;
      if (f.name.includes(":")) {
        return f.name.split(":")[1] === name;
      }
      if (f.displayName === name) return true;
      return false;
    });
    if (!field) {
      throw new Error(`Field ${name} not found`);
    }
    return field;
  }

  public getTopLevelFields(): FieldType[] {
    return this._fields;
  }

  public getAllFields(): FieldType[] {
    const fields = [];
    for (const field of this._fields) {
      fields.push(field);
      if (field.type === "object" && field.properties) {
        fields.push(...this.getFieldsRecursive(field.properties));
      }
    }
    return fields;
  }

  private getFieldsRecursive(fields: FieldType[]): FieldType[] {
    const result = [];
    for (const field of fields) {
      result.push(field);
      if (field.type === "object" && field.properties) {
        result.push(...this.getFieldsRecursive(field.properties));
      }
    }
    return result;
  }
}

export type FieldType = {
  name: string;
  type: FieldTypes;
  objectType?: string;
  defaultValue?: any;
  representedType: string;
  displayName?: string;
  properties?: FieldType[];
  primary?: boolean;
  unique?: boolean;
  notNull?: boolean;
  link?: { table: any; field: string };
  selectOptions?: { id: string; name: string }[];
  getOptionsAsync?: () => Promise<{ id: string; name: string }[]>;
  required: boolean;
  semiRequired: boolean;
};

export type FieldTypes =
  | "string"
  | "text"
  | "link"
  | "object"
  | "date"
  | "datetime"
  | "select"
  | "multiselect"
  | "boolean";

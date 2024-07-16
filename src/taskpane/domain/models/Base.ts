export class Base {
  protected _fields: FieldType[] = [];

  public getFieldByName(name: string): FieldType {
    const field = this._fields.find((f) => f.name === name);
    if (!field) {
      throw new Error(`Field ${name} not found`);
    }
    return field;
  }

  public getFields(): FieldType[] {
    return this._fields;
  }
}

export type FieldType = {
  name: string;
  type: string;
  defaultValue?: any;
  representedType: string;
  primary?: boolean;
  unique?: boolean;
  notNull?: boolean;
  link?: any;
  required: boolean;
  semiRequired: boolean;
};

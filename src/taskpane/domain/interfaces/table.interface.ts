export type TableInterface = {
  '@context': string;
  '@type': string;
  '@id': string;
  [key: string]:
    | string
    | string[]
    | {
        '@context': string;
        '@type': string;
        numerical_value: string;
        unit_of_measure: string;
      };
};

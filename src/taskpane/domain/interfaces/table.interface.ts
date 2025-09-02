export type TableInterface = {
  "@context": string | string[];
  "@type": string | string[];
  "@id": string;
  [key: string]: string | string[] | boolean | number | null | TableInterface;
};

export type TableInterface = {
  "@context": string | string[];
  "@type": string;
  "@id": string;
  [key: string]: string | string[] | boolean | TableInterface;
};

import { Address } from "./Address";
import { Characteristic } from "./Characteristic";
import { EDGProfile } from "./EDGProfile";
import { EquityDeservingGroup } from "./EquityDeservingGroup";
import { FundingState } from "./FundingState";
import { FundingStatus } from "./FundingStatus";
import { ImpactReport } from "./ImpactReport";
import { Indicator } from "./Indicator";
import { IndicatorReport } from "./IndicatorReport";
import { Organization } from "./Organization";
import { OrganizationProfile } from "./OrganizationProfile";
import { Outcome } from "./Outcome";
import { Person } from "./Person";
import { PopulationServed } from "./PopulationServed";
import { Sector } from "./Sector";
import { TeamProfile } from "./TeamProfile";
import { Theme } from "./Theme";

export const map = {
  Organization: Organization,
  Theme: Theme,
  Outcome: Outcome,
  Indicator: Indicator,
  IndicatorReport: IndicatorReport,
  Address: Address,
};

export const mapSFFModel = {
  OrganizationProfile: OrganizationProfile,
  FundingStatus: FundingStatus,
  TeamProfile: TeamProfile,
  EDGProfile: EDGProfile,
  EquityDeservingGroup: EquityDeservingGroup,
  Person: Person,
  Characteristic: Characteristic,
  FundingState: FundingState,
  Sector: Sector,
  PopulationServed: PopulationServed,
  ImpactReport: ImpactReport,
};

export type ModelType = keyof typeof map;

export type SFFModelType = keyof typeof mapSFFModel;

export function createInstance(
  sheetName: ModelType | SFFModelType
): InstanceType<(typeof map)[ModelType] | (typeof mapSFFModel)[SFFModelType]> {
  let Model: (typeof map)[ModelType] | (typeof mapSFFModel)[SFFModelType] | undefined;
  if (sheetName in map) {
    Model = map[sheetName as ModelType];
  } else if (sheetName in mapSFFModel) {
    Model = mapSFFModel[sheetName as SFFModelType];
  }
  if (!Model) {
    throw new Error(`Model for ${sheetName} not found.`);
  }
  return new Model();
}

export const ignoredFields: { [key: string]: string[] } = {
  Organization: ["hasOrganizationProfile", "hasFundingStatus", "hasImpactReport"],
  Theme: ["hasOutcome"],
  Address: ["forOrganization"],
  Person: ["forOrganizationProfile"],
  TeamProfile: ["forOrganizationProfileManagementTeam", "forOrganizationProfileBoard"],
  EquityDeservingGroup: ["forOrganizationProfile", "hasEDGProfile"],
  FundingStatus: ["forOrganizationProfile"],
  Characteristic: ["forEquityDeservingGroup"],
  EDGProfile: ["forTeamProfile"],
  FundingState: ["forFundingStatus"],
  Sector: ["forOrganizationProfile"],
  PopulationServed: ["forOrganizationProfile", "forCharacteristic"],
};

export const predefinedCodeLists = [
  "Sector",
  "PopulationServed",
  "Locality",
  "ProvinceTerritory",
  "OrganizationType",
];

export * from "./Address";
export * from "./Indicator";
export * from "./IndicatorReport";
export * from "./Organization";
export * from "./Outcome";
export * from "./Theme";

// Export SFF module classes
export * from "./Characteristic";
export * from "./EDGProfile";
export * from "./EquityDeservingGroup";
export * from "./FundingState";
export * from "./FundingStatus";
export * from "./ImpactReport";
export * from "./OrganizationProfile";
export * from "./Person";
export * from "./PopulationServed";
export * from "./Sector";
export * from "./TeamProfile";

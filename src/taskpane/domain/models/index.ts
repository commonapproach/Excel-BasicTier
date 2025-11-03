
import { Address } from "./Address";
import { Characteristic } from "./Characteristic";
import { CorporateRegistrar } from "./CorporateRegistrar";
import { EDGProfile } from "./EDGProfile";
import { EquityDeservingGroup } from "./EquityDeservingGroup";
import { FundingState } from "./FundingState";
import { FundingStatus } from "./FundingStatus";
import { Indicator } from "./Indicator";
import { IndicatorReport } from "./IndicatorReport";
import { Organization } from "./Organization";
import { OrganizationID } from "./OrganizationID";
import { OrganizationProfile } from "./OrganizationProfile";
import { Outcome } from "./Outcome";
import { Person } from "./Person";
import { Population } from "./Population";
import { PopulationServed } from "./PopulationServed";
import { ReportInfo } from "./ReportInfo";
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
	Population: Population,
};

export const mapSFFModel = {
	OrganizationProfile: OrganizationProfile,
	TeamProfile: TeamProfile,
	EDGProfile: EDGProfile,
	EquityDeservingGroup: EquityDeservingGroup,
	Person: Person,
	Characteristic: Characteristic,
	FundingStatus: FundingStatus,
	FundingState: FundingState,
	Sector: Sector,
	PopulationServed: PopulationServed,
	ReportInfo: ReportInfo,
	OrganizationID: OrganizationID,
	CorporateRegistrar: CorporateRegistrar,
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
// The order matters here
export const contextUrl = [
	"https://ontology.commonapproach.org/contexts/cidsContext.jsonld", // Base context
	"https://ontology.commonapproach.org/contexts/sffContext.jsonld", // Extended context for SFF module
];

export const ignoredFields = {
	Organization: ["hasOrganizationProfile", "hasFundingStatus", "hasReportInfo", "hasIndicatorReport", "hasID"],
	Theme: ["hasOutcome", "hasIndicator", "From field: relatesTo"],
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
	Population: ["forIndicator", "cardinalityForIndicator"],
	CorporateRegistrar: ["forOrganizationID", "issuedOrganizationID"],
	OrganizationID: ["forOrganization"],
};

export const predefinedCodeLists = [
	"Sector",
	"PopulationServed",
	"Locality",
	"ProvinceTerritory",
	"OrganizationType",
	"CorporateRegistrar",
	"EquityDeservingGroup", 
];

export * from "./Address";
export * from "./Indicator";
export * from "./IndicatorReport";
export * from "./Organization";
export * from "./Outcome";
export * from "./Population";
export * from "./Theme";

// Export SFF module classes
export * from "./Characteristic";
export * from "./CorporateRegistrar";
export * from "./EDGProfile";
export * from "./EquityDeservingGroup";
export * from "./FundingState";
export * from "./FundingStatus";
export * from "./OrganizationID";
export * from "./OrganizationProfile";
export * from "./Person";
export * from "./PopulationServed";
export * from "./ReportInfo";
export * from "./Sector";
export * from "./TeamProfile";

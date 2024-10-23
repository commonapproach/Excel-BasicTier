import {
  getAllLocalities,
  getAllOrganizationType,
  getAllProvinceTerritory,
} from "../codeLists/getCodeLists";
import { Base } from "./Base";
import { EquityDeservingGroup } from "./EquityDeservingGroup";
import { FundingStatus } from "./FundingStatus";
import { Organization } from "./Organization";
import { Person } from "./Person";
import { PopulationServed } from "./PopulationServed";
import { Sector } from "./Sector";
import { TeamProfile } from "./TeamProfile";

export class OrganizationProfile extends Base {
  static className = "OrganizationProfile";

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
        name: "forOrganization",
        type: "link",
        representedType: "string",
        defaultValue: [],
        link: { table: Organization, field: "hasOrganizationProfile" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasPrimaryContact",
        type: "link",
        representedType: "string",
        defaultValue: "",
        link: { table: Person, field: "forOrganizationProfile" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasManagementTeamProfile",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: TeamProfile, field: "forOrganizationProfileManagementTeam" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasBoardProfile",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: TeamProfile, field: "forOrganizationProfileBoard" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "sectorServed",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: Sector, field: "forOrganizationProfile" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "localityServed",
        type: "select",
        representedType: "array",
        defaultValue: [],
        selectOptions: [],
        getOptionsAsync: async () => {
          const codeList = await getAllLocalities();
          return codeList.map((item) => ({ id: item["@id"], name: item.hasName }));
        },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "provinceTerritoryServed",
        type: "select",
        representedType: "array",
        defaultValue: [],
        selectOptions: [],
        getOptionsAsync: async () => {
          const codeList = await getAllProvinceTerritory();
          return codeList.map((item) => ({ id: item["@id"], name: item.hasName }));
        },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "primaryPopulationServed",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: PopulationServed, field: "forOrganizationProfile" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "organizationType",
        type: "select",
        representedType: "array",
        defaultValue: [],
        selectOptions: [],
        getOptionsAsync: async () => {
          const codeList = await getAllOrganizationType();
          return codeList.map((item) => ({ id: item["@id"], name: item.hasName }));
        },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "servesEDG",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: EquityDeservingGroup, field: "forOrganizationProfile" },
        unique: false,
        notNull: false,
        required: false,
        semiRequired: true,
      },
      {
        name: "hasFundingStatus",
        type: "link",
        representedType: "array",
        defaultValue: [],
        link: { table: FundingStatus, field: "forOrganizationProfile" },
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
        notNull: false,
        required: false,
        semiRequired: false,
      },
    ];
  }
}

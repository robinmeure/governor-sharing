import { ISharingResult } from "./ISharingResult";
import { IContextualMenuProps } from '@fluentui/react';

export interface ISharingViewState {
  sharingLinks: ISharingResult[];
  sharingLinkIds:string[],
  sharingLinkIdsPaginated:string[],
  groups: any[];
  isOpen: boolean;
  selectedDocument: ISharingResult;
  hideSharingSettingsDialog: boolean;
  frameUrlSharingSettings:string;
  contextualMenuProps?: IContextualMenuProps;
  showResetFilters?: boolean;
  currentPage:number,
  totalPages: number,
  pageLimit: number,
  selectedTab: string,
  selectedFilter?:string,
  loadingComplete: boolean,
  statusMessage: string,
  listItems:Record<string, any>
}

export default ISharingViewState;
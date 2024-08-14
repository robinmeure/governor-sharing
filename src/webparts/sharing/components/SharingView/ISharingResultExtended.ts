import ISharingResult from "./ISharingResult";
import { SharedWith } from "./SharedWith";

export interface ISharingResultExtended {
  ptenant?:string;
  ItemType?:string;
  SnapshotDate?:string;
  SiteId?:string;
  WebId?:string;
  ItemURL?:string;
  RoleDefinition?:string;
  SharedWith?:SharedWith[];
  SharedWithCount?:SharedWith[];
  LinkId?:string;
  LinkScope?:string;
  ScopeId?:string;
  FileExtension?:string;
}
export default ISharingResultExtended;
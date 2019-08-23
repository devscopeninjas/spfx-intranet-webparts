import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDevScopeGroupsSearchProps {
  context: WebPartContext;
  itemnumberproperty: string;
  orderbyfieldproperty: string;
  ordermodefieldproperty: string;
}

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";

export interface IApprovalSystemProps {
  description: string;
  pagecontext:WebPartContext;
  spHttpClient:SPHttpClient;
  siteURL:string;
  desc:string;
  reason:string;
  currentUserName:string;
  stDate:string;
  endDate:string;
  listName:string;
  currentUserEmail:string;
}

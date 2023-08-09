import {SPHttpClient } from '@microsoft/sp-http';

 

export interface ISalesProps {
  description: string;
  siteUrl: string;
  spHttpClient:SPHttpClient;
}
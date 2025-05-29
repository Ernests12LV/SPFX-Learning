import { ISharePointList } from "../Tempalte120WebPart";

export interface ITempalte120Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  lists: ISharePointList[];

  productName: string;
  productDescription: string;
  productCost: number;
  quantity: number;
  billAmount: number;
  discount: number;
  netBillAmount: number;
  isCertified: boolean;
}

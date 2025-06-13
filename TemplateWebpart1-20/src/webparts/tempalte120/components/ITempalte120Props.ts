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
  rating: number;
  category: string;

  deliveryOption: string;
  features: string[];
  paymentMethod: string;
  colorScheme: string;

  listName: string;
  listDescription: string;
  itemTitle: string;
  itemId: string;
  listTitle: string;
  listTeam: string;
}

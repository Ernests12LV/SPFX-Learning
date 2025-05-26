import { ISharePointList } from "../Tempalte120WebPart";

export interface ITempalte120Props {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  lists: ISharePointList[]; // Add this line
}

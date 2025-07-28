import { DataProvider } from './DataProvider';
import { ISharedData } from './ISharedData';

export interface IProviderWebPartDemoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  dataProvider: DataProvider;
  onDataUpdate: (updates: Partial<ISharedData>) => void;
}

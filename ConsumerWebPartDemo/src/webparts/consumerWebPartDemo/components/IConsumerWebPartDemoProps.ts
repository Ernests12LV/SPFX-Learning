import { DynamicProperty } from '@microsoft/sp-component-base';
import { ISharedData } from './ISharedData';

export interface IConsumerWebPartDemoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  // Dynamic data properties
  sharedData: DynamicProperty<ISharedData>;
  message: DynamicProperty<string>;
  counter: DynamicProperty<number>;
  userInfo: DynamicProperty<{ displayName: string; email: string; }>;
}

export interface ISharedData {
  message: string;
  timestamp: Date;
  userInfo: {
    displayName: string;
    email: string;
  };
  counter: number;
}
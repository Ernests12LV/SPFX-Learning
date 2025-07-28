import { DynamicProperty } from '@microsoft/sp-component-base';
import { ISharedData } from './ISharedData';

export class DataProvider {
  private _data: ISharedData;
  private _onDataChanged: ((data: ISharedData) => void)[] = [];

  constructor(initialData: ISharedData) {
    this._data = initialData;
  }

  public get data(): ISharedData {
    return this._data;
  }

  public updateData(newData: Partial<ISharedData>): void {
    this._data = { ...this._data, ...newData };
    this._notifyDataChanged();
  }

  public subscribe(callback: (data: ISharedData) => void): void {
    this._onDataChanged.push(callback);
  }

  public unsubscribe(callback: (data: ISharedData) => void): void {
    const index = this._onDataChanged.indexOf(callback);
    if (index > -1) {
      this._onDataChanged.splice(index, 1);
    }
  }

  private _notifyDataChanged(): void {
    this._onDataChanged.forEach(callback => callback(this._data));
  }
}
import * as React from 'react';
import styles from './ProviderWebPartDemo.module.scss';
import type { IProviderWebPartDemoProps } from './IProviderWebPartDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISharedData } from './ISharedData';

interface IProviderWebPartDemoState {
  currentData: ISharedData;
}

export default class ProviderWebPartDemo extends React.Component<IProviderWebPartDemoProps, IProviderWebPartDemoState> {
  
  constructor(props: IProviderWebPartDemoProps) {
    super(props);
    this.state = {
      currentData: this.props.dataProvider.data
    };
  }

  public componentDidMount(): void {
    this.props.dataProvider.subscribe(this._onDataChanged);
  }

  public componentWillUnmount(): void {
    this.props.dataProvider.unsubscribe(this._onDataChanged);
  }

  private _onDataChanged = (data: ISharedData): void => {
    this.setState({ currentData: data });
  }

  private _updateMessage = (): void => {
    const newMessage = `Updated at ${new Date().toLocaleTimeString()}`;
    this.props.onDataUpdate({ 
      message: newMessage, 
      timestamp: new Date() 
    });
  }

  private _incrementCounter = (): void => {
    this.props.onDataUpdate({ 
      counter: this.state.currentData.counter + 1 
    });
  }

  private _resetData = (): void => {
    this.props.onDataUpdate({
      message: 'Reset Data',
      counter: 0,
      timestamp: new Date()
    });
  }

  public render(): React.ReactElement<IProviderWebPartDemoProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const { currentData } = this.state;

    return (
      <section className={`${styles.providerWebPartDemo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Provider Web Part Demo</h2>
          <div>Welcome, {escape(userDisplayName)}!</div>
          <div>{environmentMessage}</div>
          <div>Web part description: <strong>{escape(description)}</strong></div>
        </div>

        <div className={styles.dataSection}>
          <h3>Shared Data Provider</h3>
          <div className={styles.dataDisplay}>
            <div><strong>Current Message:</strong> {currentData.message}</div>
            <div><strong>Counter:</strong> {currentData.counter}</div>
            <div><strong>Last Updated:</strong> {currentData.timestamp.toLocaleString()}</div>
            <div><strong>User:</strong> {currentData.userInfo.displayName} ({currentData.userInfo.email})</div>
          </div>

          <div className={styles.controls}>
            <button onClick={this._updateMessage} className={styles.button}>
              Update Message
            </button>
            <button onClick={this._incrementCounter} className={styles.button}>
              Increment Counter
            </button>
            <button onClick={this._resetData} className={styles.button}>
              Reset Data
            </button>
          </div>

          <div className={styles.info}>
            <h4>How to use this Provider:</h4>
            <ol>
              <li>Add a consumer web part to the same page</li>
              <li>Configure the consumer to connect to this provider</li>
              <li>Use the buttons above to update shared data</li>
              <li>Watch the consumer web part update automatically</li>
            </ol>
          </div>
        </div>
      </section>
    );
  }
}

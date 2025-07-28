import * as React from 'react';
import styles from './ConsumerWebPartDemo.module.scss';
import type { IConsumerWebPartDemoProps } from './IConsumerWebPartDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ISharedData } from './ISharedData';

interface IConsumerWebPartDemoState {
  connectionStatus: string;
  lastUpdate: Date;
}

export default class ConsumerWebPartDemo extends React.Component<IConsumerWebPartDemoProps, IConsumerWebPartDemoState> {
  
  constructor(props: IConsumerWebPartDemoProps) {
    super(props);
    this.state = {
      connectionStatus: 'Not connected',
      lastUpdate: new Date()
    };
  }

  public componentDidMount(): void {
    this._checkConnection();
    // Subscribe to data changes
    if (this.props.sharedData) {
      this.props.sharedData.register(this.render.bind(this));
    }
    if (this.props.message) {
      this.props.message.register(this.render.bind(this));
    }
    if (this.props.counter) {
      this.props.counter.register(this.render.bind(this));
    }
    if (this.props.userInfo) {
      this.props.userInfo.register(this.render.bind(this));
    }
  }

  public componentDidUpdate(): void {
    this._checkConnection();
  }

  public componentWillUnmount(): void {
    // Unsubscribe from data changes
    if (this.props.sharedData) {
      this.props.sharedData.unregister(this.render.bind(this));
    }
    if (this.props.message) {
      this.props.message.unregister(this.render.bind(this));
    }
    if (this.props.counter) {
      this.props.counter.unregister(this.render.bind(this));
    }
    if (this.props.userInfo) {
      this.props.userInfo.unregister(this.render.bind(this));
    }
  }

  private _checkConnection(): void {
    const hasAnyConnection = this.props.sharedData?.tryGetValue() || 
                           this.props.message?.tryGetValue() || 
                           this.props.counter?.tryGetValue() || 
                           this.props.userInfo?.tryGetValue();
    
    const newStatus = hasAnyConnection ? 'Connected to Provider' : 'Not connected - Configure in Property Pane';
    
    if (newStatus !== this.state.connectionStatus) {
      this.setState({ 
        connectionStatus: newStatus,
        lastUpdate: new Date()
      });
    }
  }

  private _renderConnectionStatus(): JSX.Element {
    const isConnected = this.state.connectionStatus.includes('Connected');
    return (
      <div className={isConnected ? styles.connected : styles.disconnected}>
        <strong>Status:</strong> {this.state.connectionStatus}
        <br />
        <small>Last checked: {this.state.lastUpdate.toLocaleTimeString()}</small>
      </div>
    );
  }

  private _renderSharedData(): JSX.Element {
    const sharedData = this.props.sharedData?.tryGetValue();
    
    if (!sharedData) {
      return <div className={styles.noData}>No shared data connected</div>;
    }

    return (
      <div className={styles.dataDisplay}>
        <h4>Complete Shared Data Object:</h4>
        <div><strong>Message:</strong> {sharedData.message}</div>
        <div><strong>Counter:</strong> {sharedData.counter}</div>
        <div><strong>Timestamp:</strong> {new Date(sharedData.timestamp).toLocaleString()}</div>
        <div><strong>User:</strong> {sharedData.userInfo.displayName} ({sharedData.userInfo.email})</div>
      </div>
    );
  }

  private _renderIndividualProperties(): JSX.Element {
    const message = this.props.message?.tryGetValue();
    const counter = this.props.counter?.tryGetValue();
    const userInfo = this.props.userInfo?.tryGetValue();

    return (
      <div className={styles.individualProperties}>
        <h4>Individual Properties:</h4>
        <div className={styles.propertyItem}>
          <strong>Message Only:</strong> {message || 'Not connected'}
        </div>
        <div className={styles.propertyItem}>
          <strong>Counter Only:</strong> {counter !== undefined ? counter : 'Not connected'}
        </div>
        <div className={styles.propertyItem}>
          <strong>User Info Only:</strong> {userInfo ? `${userInfo.displayName} (${userInfo.email})` : 'Not connected'}
        </div>
      </div>
    );
  }

  public render(): React.ReactElement<IConsumerWebPartDemoProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.consumerWebPartDemo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Consumer Web Part Demo</h2>
          <div>Welcome, {escape(userDisplayName)}!</div>
          <div>{environmentMessage}</div>
          <div>Web part description: <strong>{escape(description)}</strong></div>
        </div>

        <div className={styles.consumerContent}>
          <h3>Dynamic Data Consumer</h3>
          
          {this._renderConnectionStatus()}
          
          <div className={styles.dataSection}>
            {this._renderSharedData()}
          </div>

          <div className={styles.dataSection}>
            {this._renderIndividualProperties()}
          </div>

          <div className={styles.instructions}>
            <h4>How to Connect:</h4>
            <ol>
              <li>Make sure the Provider Web Part is added to the same page</li>
              <li>Edit this web part and go to the Property Pane</li>
              <li>In the "Dynamic Data Connection" section, connect to the provider properties</li>
              <li>Save the configuration</li>
              <li>This web part will automatically update when provider data changes</li>
            </ol>
          </div>
        </div>
      </section>
    );
  }
}

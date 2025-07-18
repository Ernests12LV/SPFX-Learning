import * as React from 'react';
import styles from './GraphApiDemo2React.module.scss';
import type { IGraphApiDemo2ReactProps } from './IGraphApiDemo2ReactProps';
import { escape } from '@microsoft/sp-lodash-subset';
//import { Client } from '@microsoft/microsoft-graph-client';
//import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IEvent {
  id: string;
  subject: string;
  start: { dateTime: string };
  end: { dateTime: string };
  organizer: { emailAddress: { name: string } };
}

export default class GraphApiDemo2React extends React.Component<IGraphApiDemo2ReactProps, { events: IEvent[] }> {
  constructor(props: IGraphApiDemo2ReactProps) {
    super(props);
    this.state = { events: [] };
  }

  public componentDidMount(): void {
    this._getEvents();
  }

  private async _getEvents() {
    const client = await this.props.context.msGraphClientFactory.getClient('3');
    const result = await client
      .api('/me/events')
      .select('subject,start,end,organizer')
      .orderby('start/dateTime')
      .top(10)
      .get();

    this.setState({ events: result.value });
  }

  public render(): React.ReactElement<IGraphApiDemo2ReactProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.graphApiDemo2React} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>
            The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
          </p>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
            <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
            <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
          </ul>
        </div>
        <div>
          <h3>Your Upcoming Outlook Events</h3>
          <ul>
            {this.state.events.map(ev => (
              <li key={ev.id}>
                <strong>{ev.subject || "(No subject)"}</strong><br />
                {new Date(ev.start.dateTime).toLocaleString()} - {new Date(ev.end.dateTime).toLocaleString()}<br />
                Organizer: {ev.organizer.emailAddress.name}
              </li>
            ))}
            {this.state.events.length === 0 && <li>No events found.</li>}
          </ul>
        </div>
      </section>
    );
  }
}

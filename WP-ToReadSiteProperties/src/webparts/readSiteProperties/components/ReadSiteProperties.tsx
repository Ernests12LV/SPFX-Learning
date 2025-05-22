import * as React from 'react';
import styles from './ReadSiteProperties.module.scss';
import type { IReadSitePropertiesProps } from './IReadSitePropertiesProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ReadSiteProperties extends React.Component<IReadSitePropertiesProps> {
  public render(): React.ReactElement<IReadSitePropertiesProps> {
    const {
      environemtTitle,
      environment,
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      absoluteUrl,
      siteTitle,
      relativeUrl
    } = this.props;

    return (
      <section className={`${styles.readSiteProperties} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>

          <p>Absolute URL{escape(absoluteUrl)}</p>
          <p>Site Title: {escape(siteTitle)}</p>
          <p>Relative URL: {escape(relativeUrl)}</p>
          <p>User Name {escape(userDisplayName)}</p>

          <p>Environment {environment}</p>
          <p>Environment Title: {environemtTitle}</p>

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
      </section>
    );
  }
}

import * as React from 'react';
import styles from './CultureInfoDetails.module.scss';
import type { ICultureInfoDetailsProps } from './CultureInfoDetailsProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CultureInfoDetails extends React.Component<ICultureInfoDetailsProps> {
  public render(): React.ReactElement<ICultureInfoDetailsProps> {
    const {
      cultureName,
      uiCultureName,
      isRightToLeft,
      isDarkTheme,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.CultureInfoDetails} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
            <li>current Culture Name {}</li>
          </ul>
        </div>
      </section>
    );
  }
}

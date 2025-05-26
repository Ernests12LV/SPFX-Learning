import * as React from 'react';
import styles from './Tempalte120.module.scss';
import type { ITempalte120Props } from './ITempalte120Props';
import { escape } from '@microsoft/sp-lodash-subset';

export default class Tempalte120 extends React.Component<ITempalte120Props> {
  
  public render(): React.ReactElement<ITempalte120Props> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      lists
    } = this.props;

    return (
      <section className={`${styles.tempalte120} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
        <div>
          <h4>Learn more about SPFx development:</h4>
          <ul className={styles.links}>
            <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
          </ul>
        </div>
        <div>
          <h4>SharePoint Lists:</h4>
          <ul>
            {lists.map((list) => (
              <li key={list.ListId}>
                {list.ListTitle} ({list.ListId})
              </li>
            ))}
          </ul>
        </div>
      </section>
    );
  }
}

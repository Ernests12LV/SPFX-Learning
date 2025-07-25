import * as React from 'react';
import styles from './GraphApiDemo2React.module.scss';
import type { IGraphApiDemo2ReactProps } from './IGraphApiDemo2ReactProps';
import { escape } from '@microsoft/sp-lodash-subset';

interface IEvent {
  id: string;
  subject: string;
  start: { dateTime: string };
  end: { dateTime: string };
  organizer: { emailAddress: { name: string } };
}

interface IUser {
  id: string;
  displayName: string;
  mail: string;
  jobTitle: string;
  department: string;
  officeLocation: string;
}

interface IGraphApiDemo2ReactState {
  events: IEvent[];
  searchEmail: string;
  searchedUser: IUser | null;
  isSearching: boolean;
  searchError: string;
}

export default class GraphApiDemo2React extends React.Component<IGraphApiDemo2ReactProps, IGraphApiDemo2ReactState> {
  constructor(props: IGraphApiDemo2ReactProps) {
    super(props);
    this.state = { 
      events: [],
      searchEmail: '',
      searchedUser: null,
      isSearching: false,
      searchError: ''
    };
  }

  public componentDidMount(): void {
    this._getEvents();
  }

  private async _getEvents() {
    try {
      const client = await this.props.context.msGraphClientFactory.getClient('3');
      const result = await client
        .api('/me/events')
        .select('subject,start,end,organizer')
        .orderby('start/dateTime')
        .top(10)
        .get();

      this.setState({ events: result.value });
    } catch (error) {
      console.error('Error fetching events:', error);
    }
  }

  private async _searchUser() {
    if (!this.state.searchEmail.trim()) {
      this.setState({ searchError: 'Please enter an email address' });
      return;
    }

    this.setState({ isSearching: true, searchError: '', searchedUser: null });

    try {
      const client = await this.props.context.msGraphClientFactory.getClient('3');
      
      // Search for user by email
      const result = await client
        .api('/users')
        .filter(`mail eq '${this.state.searchEmail}' or userPrincipalName eq '${this.state.searchEmail}'`)
        .select('id,displayName,mail,jobTitle,department,officeLocation')
        .get();

      if (result.value && result.value.length > 0) {
        this.setState({ 
          searchedUser: result.value[0],
          isSearching: false,
          searchError: ''
        });
      } else {
        this.setState({ 
          searchedUser: null,
          isSearching: false,
          searchError: 'User not found'
        });
      }
    } catch (error) {
      console.error('Error searching user:', error);
      this.setState({ 
        isSearching: false,
        searchError: 'Error searching for user. Please check permissions.'
      });
    }
  }

  private _onEmailChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    this.setState({ searchEmail: event.target.value });
  }

  private _onSearchClick = () => {
    this._searchUser();
  }

  private _onKeyPress = (event: React.KeyboardEvent<HTMLInputElement>) => {
    if (event.key === 'Enter') {
      this._searchUser();
    }
  }

  public render(): React.ReactElement<IGraphApiDemo2ReactProps> {
    const {
      // description,
      // isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.graphApiDemo2React} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <h2>Welcome, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
        </div>

        {/* Calendar Events Section */}
        <div className={styles.section}>
          <h3>Your Upcoming Outlook Events</h3>
          <ul className={styles.eventsList}>
            {this.state.events.map(ev => (
              <li key={ev.id} className={styles.eventItem}>
                <strong>{ev.subject || "(No subject)"}</strong><br />
                <span className={styles.eventTime}>
                  {new Date(ev.start.dateTime).toLocaleString()} - {new Date(ev.end.dateTime).toLocaleString()}
                </span><br />
                <span className={styles.organizer}>
                  Organizer: {ev.organizer.emailAddress.name}
                </span>
              </li>
            ))}
            {this.state.events.length === 0 && <li>No events found.</li>}
          </ul>
        </div>

        {/* User Search Section */}
        <div className={styles.section}>
          <h3>Search User by Email</h3>
          <div className={styles.searchContainer}>
            <input
              type="email"
              placeholder="Enter user email address"
              value={this.state.searchEmail}
              onChange={this._onEmailChange}
              onKeyPress={this._onKeyPress}
              className={styles.searchInput}
            />
            <button
              onClick={this._onSearchClick}
              disabled={this.state.isSearching}
              className={styles.searchButton}
            >
              {this.state.isSearching ? 'Searching...' : 'Search'}
            </button>
          </div>

          {this.state.searchError && (
            <div className={styles.errorMessage}>
              {this.state.searchError}
            </div>
          )}

          {this.state.searchedUser && (
            <div className={styles.userCard}>
              <h4>User Found:</h4>
              <div className={styles.userInfo}>
                <div><strong>Name:</strong> {this.state.searchedUser.displayName}</div>
                <div><strong>Email:</strong> {this.state.searchedUser.mail}</div>
                <div><strong>Job Title:</strong> {this.state.searchedUser.jobTitle || 'N/A'}</div>
                <div><strong>Department:</strong> {this.state.searchedUser.department || 'N/A'}</div>
                <div><strong>Office Location:</strong> {this.state.searchedUser.officeLocation || 'N/A'}</div>
              </div>
            </div>
          )}
        </div>
      </section>
    );
  }
}

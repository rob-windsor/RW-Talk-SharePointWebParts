import * as React from 'react';
import styles from './HelloWebPartsReact.module.scss';
import type { IHelloWebPartsReactProps } from './IHelloWebPartsReactProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class HelloWebPartsReact extends React.Component<IHelloWebPartsReactProps> {
  public render(): React.ReactElement<IHelloWebPartsReactProps> {
    const {
      spListItems,
      description,
      listName,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.helloWebPartsReact} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Description property value: <strong>{escape(description)}</strong></div>
          <div>List name property value: <strong>{escape(listName)}</strong></div>
        </div>
        <div className={styles.mySection}>
          <button type="button" onClick={this.onWelcomeButtonClicked}>Show Welcome Message</button>
        </div>
        <div className={styles.mySection}>
          <button type="button" onClick={this.onGetListItemsClicked}>Get List Items</button>
        </div>
        <div className={styles.mySection}>
          <ul>
            {spListItems && spListItems.map((listItem) =>
              <li key={listItem.Id}>
                <strong>Id:</strong> {listItem.Id}, <strong>Title:</strong> {listItem.Title}
              </li>
            )
            }
          </ul>
        </div>        
      </section>
    );
  }

  private onWelcomeButtonClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    alert('Welcome to SharePoint Framework!');
  }

  private onGetListItemsClicked = (event: React.MouseEvent<HTMLButtonElement>): void => {
    event.preventDefault();
    if (this.props.onGetListItems) this.props.onGetListItems();
  }
}

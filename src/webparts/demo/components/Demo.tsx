import * as React from 'react';
import styles from './Demo.module.scss';
import type { IDemoProps } from './IDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  Pivot,
  PivotItem,
  Stack,
  Text
} from '@fluentui/react';
import { SharePointService, IListInfo, IUserInfo, ISiteInfo } from '../services/SharePointService';
import { IListItem } from '../models/IListItem';
import { ListViewer } from './ListViewer';
import { ItemViewer } from './ItemViewer';
import { UserInfo } from './UserInfo';

export interface IDemoState {
  // Lists and items
  lists: IListInfo[];
  selectedListItems: IListItem[];
  selectedListTitle: string;
  
  // User and site info
  userInfo?: IUserInfo;
  siteInfo?: ISiteInfo;
  
  // Loading states
  listsLoading: boolean;
  itemsLoading: boolean;
  userInfoLoading: boolean;
  
  // Error states
  listsError?: string;
  itemsError?: string;
  userInfoError?: string;
  
  // UI state
  currentView: 'welcome' | 'lists' | 'items' | 'profile';
}

export default class Demo extends React.Component<IDemoProps, IDemoState> {
  private _sharePointService: SharePointService;

  constructor(props: IDemoProps) {
    super(props);
    
    this._sharePointService = new SharePointService(this.props.context);
    
    this.state = {
      lists: [],
      selectedListItems: [],
      selectedListTitle: '',
      listsLoading: false,
      itemsLoading: false,
      userInfoLoading: false,
      currentView: 'welcome'
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._loadUserAndSiteInfo();
  }

  private _loadUserAndSiteInfo = async (): Promise<void> => {
    this.setState({ userInfoLoading: true, userInfoError: undefined });
    
    try {
      const [userInfo, siteInfo] = await Promise.all([
        this._sharePointService.getCurrentUser(),
        this._sharePointService.getSiteInfo()
      ]);
      
      this.setState({
        userInfo,
        siteInfo,
        userInfoLoading: false
      });
    } catch (error) {
      this.setState({
        userInfoError: error.message || 'Failed to load user and site information',
        userInfoLoading: false
      });
    }
  };

  private _loadLists = async (): Promise<void> => {
    this.setState({ listsLoading: true, listsError: undefined });
    
    try {
      const lists = await this._sharePointService.getLists();
      this.setState({
        lists,
        listsLoading: false,
        currentView: 'lists'
      });
    } catch (error) {
      this.setState({
        listsError: error.message || 'Failed to load lists',
        listsLoading: false
      });
    }
  };

  private _loadListItems = async (listTitle: string): Promise<void> => {
    this.setState({ 
      itemsLoading: true, 
      itemsError: undefined,
      selectedListTitle: listTitle 
    });
    
    try {
      const items = await this._sharePointService.getListItems(listTitle, 50);
      this.setState({
        selectedListItems: items,
        itemsLoading: false,
        currentView: 'items'
      });
    } catch (error) {
      this.setState({
        itemsError: error.message || 'Failed to load list items',
        itemsLoading: false
      });
    }
  };

  private _createListItem = async (itemData: { Title: string; Description?: string }): Promise<void> => {
    try {
      const success = await this._sharePointService.createListItem(
        this.state.selectedListTitle,
        itemData
      );
      
      if (success) {
        // Reload items to show the new item
        await this._loadListItems(this.state.selectedListTitle);
      } else {
        throw new Error('Failed to create item');
      }
    } catch (error) {
      console.error('Error creating item:', error);
      throw error;
    }
  };

  private _onBackToLists = (): void => {
    this.setState({
      currentView: 'lists',
      selectedListItems: [],
      selectedListTitle: '',
      itemsError: undefined
    });
  };

  private _renderWelcomeView = (): JSX.Element => {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.demo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img 
            alt="" 
            src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} 
            className={styles.welcomeImage} 
          />
          <h2>Welcome, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property: <strong>{escape(description)}</strong></div>
        </div>
        
        <Stack tokens={{ childrenGap: 20 }}>
          <div>
            <h3>SharePoint Framework Demo Features</h3>
            <p>
              This comprehensive SPFx demo showcases various SharePoint integration capabilities:
            </p>
            
            <ul className={styles.links}>
              <li><strong>Lists Management:</strong> View and interact with SharePoint lists</li>
              <li><strong>CRUD Operations:</strong> Create, read, update, and delete list items</li>
              <li><strong>User Profile:</strong> Display current user and site information</li>
              <li><strong>Modern UI:</strong> Built with Fluent UI React components</li>
              <li><strong>Responsive Design:</strong> Works on desktop, tablet, and mobile devices</li>
            </ul>
          </div>

          <div>
            <h4>Learn more about SPFx development:</h4>
            <ul className={styles.links}>
              <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
              <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
              <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
              <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
            </ul>
          </div>
        </Stack>
      </section>
    );
  };

  private _renderCurrentView = (): JSX.Element => {
    const { currentView, lists, selectedListItems, selectedListTitle, 
            listsLoading, itemsLoading, listsError, itemsError,
            userInfo, siteInfo, userInfoLoading, userInfoError } = this.state;

    switch (currentView) {
      case 'lists':
        return (
          <ListViewer
            lists={lists}
            loading={listsLoading}
            error={listsError}
            onRefresh={this._loadLists}
            onListSelect={this._loadListItems}
          />
        );
      
      case 'items':
        return (
          <ItemViewer
            listTitle={selectedListTitle}
            items={selectedListItems}
            loading={itemsLoading}
            error={itemsError}
            onRefresh={() => this._loadListItems(selectedListTitle)}
            onBack={this._onBackToLists}
            onCreateItem={this._createListItem}
          />
        );
      
      case 'profile':
        return (
          <UserInfo
            userInfo={userInfo}
            siteInfo={siteInfo}
            loading={userInfoLoading}
            error={userInfoError}
          />
        );
      
      default:
        return this._renderWelcomeView();
    }
  };

  public render(): React.ReactElement<IDemoProps> {
    return (
      <div className={styles.demo}>
        <Stack tokens={{ childrenGap: 20 }}>
          <Text variant="xxLarge" style={{ fontWeight: 'bold', textAlign: 'center' }}>
            SharePoint Framework Demo
          </Text>
          
          <Pivot
            selectedKey={this.state.currentView}
            onLinkClick={(item) => {
              const key = item?.props.itemKey as IDemoState['currentView'];
              if (key === 'lists' && this.state.lists.length === 0) {
                this._loadLists().catch((error) => {
                  console.error('Error loading lists:', error);
                });
              } else {
                this.setState({ currentView: key });
              }
            }}
          >
            <PivotItem headerText="Welcome" itemKey="welcome" />
            <PivotItem headerText="Lists" itemKey="lists" />
            <PivotItem headerText="User Profile" itemKey="profile" />
          </Pivot>

          {this._renderCurrentView()}
        </Stack>
      </div>
    );
  }
}

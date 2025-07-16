import * as React from 'react';
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  CommandBar,
  ICommandBarItemProps,
  SearchBox,
  Stack
} from '@fluentui/react';
import { IListInfo } from '../services/SharePointService';
import styles from './Demo.module.scss';

export interface IListViewerProps {
  lists: IListInfo[];
  loading: boolean;
  error?: string;
  onRefresh: () => void;
  onListSelect: (listTitle: string) => void;
}

export interface IListViewerState {
  filteredLists: IListInfo[];
  searchQuery: string;
}

export class ListViewer extends React.Component<IListViewerProps, IListViewerState> {
  private _columns: IColumn[];

  constructor(props: IListViewerProps) {
    super(props);

    this.state = {
      filteredLists: this.props.lists,
      searchQuery: ''
    };

    this._columns = [
      {
        key: 'title',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 150,
        maxWidth: 250,
        isResizable: true,
        onRender: (item: IListInfo) => (
          <span 
            style={{ cursor: 'pointer', color: '#0078d4' }}
            onClick={() => this.props.onListSelect(item.Title)}
          >
            {item.Title}
          </span>
        )
      },
      {
        key: 'itemCount',
        name: 'Item Count',
        fieldName: 'ItemCount',
        minWidth: 100,
        maxWidth: 150,
        isResizable: true
      },
      {
        key: 'created',
        name: 'Created',
        fieldName: 'Created',
        minWidth: 150,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IListInfo) => (
          <span>{new Date(item.Created).toLocaleDateString()}</span>
        )
      },
      {
        key: 'baseType',
        name: 'Type',
        fieldName: 'BaseType',
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        onRender: (item: IListInfo) => (
          <span>{item.BaseType === 0 ? 'List' : 'Document Library'}</span>
        )
      }
    ];
  }

  public componentDidUpdate(prevProps: IListViewerProps): void {
    if (prevProps.lists !== this.props.lists) {
      this._filterLists(this.state.searchQuery);
    }
  }

  private _onSearch = (newValue?: string): void => {
    const searchQuery = newValue || '';
    this.setState({ searchQuery });
    this._filterLists(searchQuery);
  };

  private _filterLists = (query: string): void => {
    const filteredLists = query
      ? this.props.lists.filter(list => 
          list.Title.toLowerCase().indexOf(query.toLowerCase()) !== -1
        )
      : this.props.lists;
    
    this.setState({ filteredLists });
  };

  private _getCommandBarItems = (): ICommandBarItemProps[] => {
    return [
      {
        key: 'refresh',
        text: 'Refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: this.props.onRefresh
      }
    ];
  };

  public render(): React.ReactElement<IListViewerProps> {
    const { loading, error } = this.props;
    const { filteredLists } = this.state;

    if (loading) {
      return (
        <div className={styles.center}>
          <Spinner size={SpinnerSize.large} label="Loading lists..." />
        </div>
      );
    }

    if (error) {
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          Error loading lists: {error}
        </MessageBar>
      );
    }

    return (
      <div className={styles.listViewer}>
        <Stack tokens={{ childrenGap: 15 }}>
          <h3>SharePoint Lists</h3>
          
          <CommandBar items={this._getCommandBarItems()} />
          
          <SearchBox
            placeholder="Search lists..."
            onSearch={this._onSearch}
            onClear={() => this._onSearch('')}
          />

          {filteredLists.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No lists found matching your search criteria.
            </MessageBar>
          ) : (
            <DetailsList
              items={filteredLists}
              columns={this._columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
              className={styles.detailsList}
            />
          )}
        </Stack>
      </div>
    );
  }
}

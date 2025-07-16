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
  TextField,
  PrimaryButton,
  DefaultButton,
  Stack,
  Panel,
  PanelType
} from '@fluentui/react';
import { IListItem } from '../models/IListItem';
import styles from './Demo.module.scss';

export interface IItemViewerProps {
  listTitle: string;
  items: IListItem[];
  loading: boolean;
  error?: string;
  onRefresh: () => void;
  onBack: () => void;
  onCreateItem: (itemData: { Title: string; Description?: string }) => void;
}

export interface IItemViewerState {
  isCreatePanelOpen: boolean;
  newItemTitle: string;
  newItemDescription: string;
  creatingItem: boolean;
}

export class ItemViewer extends React.Component<IItemViewerProps, IItemViewerState> {
  private _columns: IColumn[];

  constructor(props: IItemViewerProps) {
    super(props);

    this.state = {
      isCreatePanelOpen: false,
      newItemTitle: '',
      newItemDescription: '',
      creatingItem: false
    };

    this._columns = [
      {
        key: 'id',
        name: 'ID',
        fieldName: 'Id',
        minWidth: 50,
        maxWidth: 80,
        isResizable: true
      },
      {
        key: 'title',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 150,
        maxWidth: 300,
        isResizable: true
      },
      {
        key: 'author',
        name: 'Author',
        fieldName: 'Author',
        minWidth: 120,
        maxWidth: 200,
        isResizable: true,
        onRender: (item: IListItem) => (
          <span>{item.Author.Title}</span>
        )
      },
      {
        key: 'created',
        name: 'Created',
        fieldName: 'Created',
        minWidth: 120,
        maxWidth: 180,
        isResizable: true,
        onRender: (item: IListItem) => (
          <span>{new Date(item.Created).toLocaleDateString()}</span>
        )
      },
      {
        key: 'modified',
        name: 'Modified',
        fieldName: 'Modified',
        minWidth: 120,
        maxWidth: 180,
        isResizable: true,
        onRender: (item: IListItem) => (
          <span>{new Date(item.Modified).toLocaleDateString()}</span>
        )
      }
    ];
  }

  private _getCommandBarItems = (): ICommandBarItemProps[] => {
    return [
      {
        key: 'back',
        text: 'Back to Lists',
        iconProps: { iconName: 'Back' },
        onClick: this.props.onBack
      },
      {
        key: 'refresh',
        text: 'Refresh',
        iconProps: { iconName: 'Refresh' },
        onClick: this.props.onRefresh
      },
      {
        key: 'create',
        text: 'New Item',
        iconProps: { iconName: 'Add' },
        onClick: () => this.setState({ isCreatePanelOpen: true })
      }
    ];
  };

  private _onCreateItem = async (): Promise<void> => {
    const { newItemTitle, newItemDescription } = this.state;
    
    if (!newItemTitle.trim()) {
      return;
    }

    this.setState({ creatingItem: true });

    try {
      await this.props.onCreateItem({
        Title: newItemTitle.trim(),
        Description: newItemDescription.trim() || undefined
      });

      this.setState({
        isCreatePanelOpen: false,
        newItemTitle: '',
        newItemDescription: '',
        creatingItem: false
      });
    } catch (error) {
      console.error('Error creating item:', error);
      this.setState({ creatingItem: false });
    }
  };

  private _onCancelCreate = (): void => {
    this.setState({
      isCreatePanelOpen: false,
      newItemTitle: '',
      newItemDescription: '',
      creatingItem: false
    });
  };

  public render(): React.ReactElement<IItemViewerProps> {
    const { listTitle, items, loading, error } = this.props;
    const { isCreatePanelOpen, newItemTitle, newItemDescription, creatingItem } = this.state;

    if (loading) {
      return (
        <div className={styles.center}>
          <Spinner size={SpinnerSize.large} label={`Loading items from ${listTitle}...`} />
        </div>
      );
    }

    if (error) {
      return (
        <MessageBar messageBarType={MessageBarType.error}>
          Error loading items: {error}
        </MessageBar>
      );
    }

    return (
      <div className={styles.itemViewer}>
        <Stack tokens={{ childrenGap: 15 }}>
          <h3>Items in &ldquo;{listTitle}&rdquo; List</h3>
          
          <CommandBar items={this._getCommandBarItems()} />

          {items.length === 0 ? (
            <MessageBar messageBarType={MessageBarType.info}>
              No items found in this list.
            </MessageBar>
          ) : (
            <DetailsList
              items={items}
              columns={this._columns}
              setKey="set"
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
              isHeaderVisible={true}
              className={styles.detailsList}
            />
          )}
        </Stack>

        <Panel
          headerText="Create New Item"
          isOpen={isCreatePanelOpen}
          onDismiss={this._onCancelCreate}
          type={PanelType.medium}
          closeButtonAriaLabel="Close"
        >
          <div className={styles.createForm}>
            <Stack tokens={{ childrenGap: 15 }}>
              <TextField
                label="Title *"
                value={newItemTitle}
                onChange={(_, newValue) => this.setState({ newItemTitle: newValue || '' })}
                required
                placeholder="Enter item title"
              />
              
              <TextField
                label="Description"
                value={newItemDescription}
                onChange={(_, newValue) => this.setState({ newItemDescription: newValue || '' })}
                multiline
                rows={4}
                placeholder="Enter item description (optional)"
              />

              <Stack horizontal tokens={{ childrenGap: 10 }}>
                <PrimaryButton
                  text="Create Item"
                  onClick={this._onCreateItem}
                  disabled={!newItemTitle.trim() || creatingItem}
                />
                <DefaultButton
                  text="Cancel"
                  onClick={this._onCancelCreate}
                  disabled={creatingItem}
                />
              </Stack>

              {creatingItem && (
                <Spinner size={SpinnerSize.small} label="Creating item..." />
              )}
            </Stack>
          </div>
        </Panel>
      </div>
    );
  }
}

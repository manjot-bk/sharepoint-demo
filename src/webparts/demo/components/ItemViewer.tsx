import * as React from 'react'
import { useState, useMemo, useCallback } from 'react'
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
	PanelType,
} from '@fluentui/react'
import { IListItem } from '../models/IListItem'
import styles from './Demo.module.scss'

export interface IItemViewerProps {
	listTitle: string
	items: IListItem[]
	loading: boolean
	error?: string
	onRefresh: () => void
	onBack: () => void
	onCreateItem: (itemData: { Title: string; Description?: string }) => void
}

export const ItemViewer: React.FunctionComponent<IItemViewerProps> = (props) => {
	const [isCreatePanelOpen, setIsCreatePanelOpen] = useState<boolean>(false)
	const [newItemTitle, setNewItemTitle] = useState<string>('')
	const [newItemDescription, setNewItemDescription] = useState<string>('')
	const [creatingItem, setCreatingItem] = useState<boolean>(false)

	// Define columns with memoization
	const columns: IColumn[] = useMemo(
		() => [
			{
				key: 'id',
				name: 'ID',
				fieldName: 'Id',
				minWidth: 50,
				maxWidth: 80,
				isResizable: true,
			},
			{
				key: 'title',
				name: 'Title',
				fieldName: 'Title',
				minWidth: 150,
				maxWidth: 300,
				isResizable: true,
			},
			{
				key: 'author',
				name: 'Author',
				fieldName: 'Author',
				minWidth: 120,
				maxWidth: 200,
				isResizable: true,
				onRender: (item: IListItem) => <span>{item.Author.Title}</span>,
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
				),
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
				),
			},
		],
		[]
	)

	// Command bar items
	const commandBarItems: ICommandBarItemProps[] = useMemo(
		() => [
			{
				key: 'back',
				text: 'Back to Lists',
				iconProps: { iconName: 'Back' },
				onClick: props.onBack,
			},
			{
				key: 'refresh',
				text: 'Refresh',
				iconProps: { iconName: 'Refresh' },
				onClick: props.onRefresh,
			},
			{
				key: 'create',
				text: 'New Item',
				iconProps: { iconName: 'Add' },
				onClick: () => setIsCreatePanelOpen(true),
			},
		],
		[props.onBack, props.onRefresh]
	)

	// Handle create item
	const onCreateItem = useCallback(async (): Promise<void> => {
		if (!newItemTitle.trim()) {
			return
		}

		setCreatingItem(true)

		try {
			await props.onCreateItem({
				Title: newItemTitle.trim(),
				Description: newItemDescription.trim() || undefined,
			})

			setIsCreatePanelOpen(false)
			setNewItemTitle('')
			setNewItemDescription('')
			setCreatingItem(false)
		} catch (error) {
			console.error('Error creating item:', error)
			setCreatingItem(false)
		}
	}, [newItemTitle, newItemDescription, props.onCreateItem])

	// Handle cancel create
	const onCancelCreate = useCallback((): void => {
		setIsCreatePanelOpen(false)
		setNewItemTitle('')
		setNewItemDescription('')
		setCreatingItem(false)
	}, [])

	// Loading state
	if (props.loading) {
		return (
			<div className={styles.center}>
				<Spinner
					size={SpinnerSize.large}
					label={`Loading items from ${props.listTitle}...`}
				/>
			</div>
		)
	}

	// Error state
	if (props.error) {
		return (
			<MessageBar messageBarType={MessageBarType.error}>
				Error loading items: {props.error}
			</MessageBar>
		)
	}

	// Main render
	return (
		<div className={styles.itemViewer}>
			<Stack tokens={{ childrenGap: 15 }}>
				<h3>Items in &ldquo;{props.listTitle}&rdquo; List</h3>

				<CommandBar items={commandBarItems} />

				{props.items.length === 0 ? (
					<MessageBar messageBarType={MessageBarType.info}>
						No items found in this list.
					</MessageBar>
				) : (
					<DetailsList
						items={props.items}
						columns={columns}
						setKey='set'
						layoutMode={DetailsListLayoutMode.justified}
						selectionMode={SelectionMode.none}
						isHeaderVisible={true}
						className={styles.detailsList}
					/>
				)}
			</Stack>

			<Panel
				headerText='Create New Item'
				isOpen={isCreatePanelOpen}
				onDismiss={onCancelCreate}
				type={PanelType.medium}
				closeButtonAriaLabel='Close'
			>
				<div className={styles.createForm}>
					<Stack tokens={{ childrenGap: 15 }}>
						<TextField
							label='Title *'
							value={newItemTitle}
							onChange={(_, newValue) => setNewItemTitle(newValue || '')}
							required
							placeholder='Enter item title'
						/>

						<TextField
							label='Description'
							value={newItemDescription}
							onChange={(_, newValue) => setNewItemDescription(newValue || '')}
							multiline
							rows={4}
							placeholder='Enter item description (optional)'
						/>

						<Stack horizontal tokens={{ childrenGap: 10 }}>
							<PrimaryButton
								text='Create Item'
								onClick={onCreateItem}
								disabled={!newItemTitle.trim() || creatingItem}
							/>
							<DefaultButton
								text='Cancel'
								onClick={onCancelCreate}
								disabled={creatingItem}
							/>
						</Stack>

						{creatingItem && (
							<Spinner size={SpinnerSize.small} label='Creating item...' />
						)}
					</Stack>
				</div>
			</Panel>
		</div>
	)
}

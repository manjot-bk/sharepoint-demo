import * as React from 'react'
import { useState, useEffect, useMemo, useCallback } from 'react'
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
	Stack,
} from '@fluentui/react'
import { IListInfo } from '../services/SharePointService'
import styles from './Demo.module.scss'

export interface IListViewerProps {
	lists: IListInfo[]
	loading: boolean
	error?: string
	onRefresh: () => void
	onListSelect: (listTitle: string) => void
}

export const ListViewer: React.FunctionComponent<IListViewerProps> = (props) => {
	const [filteredLists, setFilteredLists] = useState<IListInfo[]>(props.lists)
	const [searchQuery, setSearchQuery] = useState<string>('')

	// Define columns with memoization for performance
	const columns: IColumn[] = useMemo(
		() => [
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
						onClick={() => props.onListSelect(item.Title)}
					>
						{item.Title}
					</span>
				),
			},
			{
				key: 'itemCount',
				name: 'Item Count',
				fieldName: 'ItemCount',
				minWidth: 100,
				maxWidth: 150,
				isResizable: true,
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
				),
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
				),
			},
		],
		[props.onListSelect]
	)

	// Filter lists based on search query
	const filterLists = useCallback(
		(query: string): void => {
			const filtered = query
				? props.lists.filter(
						(list) => list.Title.toLowerCase().indexOf(query.toLowerCase()) !== -1
				  )
				: props.lists

			setFilteredLists(filtered)
		},
		[props.lists]
	)

	// Handle search
	const onSearch = useCallback(
		(newValue?: string): void => {
			const searchValue = newValue || ''
			setSearchQuery(searchValue)
			filterLists(searchValue)
		},
		[filterLists]
	)

	// Command bar items
	const commandBarItems: ICommandBarItemProps[] = useMemo(
		() => [
			{
				key: 'refresh',
				text: 'Refresh',
				iconProps: { iconName: 'Refresh' },
				onClick: props.onRefresh,
			},
		],
		[props.onRefresh]
	)

	// Update filtered lists when props change
	useEffect(() => {
		filterLists(searchQuery)
	}, [props.lists, searchQuery, filterLists])

	// Loading state
	if (props.loading) {
		return (
			<div className={styles.center}>
				<Spinner size={SpinnerSize.large} label='Loading lists...' />
			</div>
		)
	}

	// Error state
	if (props.error) {
		return (
			<MessageBar messageBarType={MessageBarType.error}>
				Error loading lists: {props.error}
			</MessageBar>
		)
	}

	// Main render
	return (
		<div className={styles.listViewer}>
			<Stack tokens={{ childrenGap: 15 }}>
				<h3>SharePoint Lists</h3>

				<CommandBar items={commandBarItems} />

				<SearchBox
					placeholder='Search lists...'
					onSearch={onSearch}
					onClear={() => onSearch('')}
				/>

				{filteredLists.length === 0 ? (
					<MessageBar messageBarType={MessageBarType.info}>
						No lists found matching your search criteria.
					</MessageBar>
				) : (
					<DetailsList
						items={filteredLists}
						columns={columns}
						setKey='set'
						layoutMode={DetailsListLayoutMode.justified}
						selectionMode={SelectionMode.none}
						isHeaderVisible={true}
						className={styles.detailsList}
					/>
				)}
			</Stack>
		</div>
	)
}

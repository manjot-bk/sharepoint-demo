import * as React from 'react'
import { useState, useEffect, useCallback, useMemo } from 'react'
import styles from './Demo.module.scss'
import type { IDemoProps } from './IDemoProps'
import { escape } from '@microsoft/sp-lodash-subset'
import { Pivot, PivotItem, Stack, Text } from '@fluentui/react'
import {
	SharePointService,
	IListInfo,
	IUserInfo,
	ISiteInfo,
} from '../services/SharePointService'
import { IListItem } from '../models/IListItem'
import { ListViewer } from './ListViewer'
import { ItemViewer } from './ItemViewer'
import { UserInfo } from './UserInfo'

type CurrentView = 'welcome' | 'lists' | 'items' | 'profile'

const Demo: React.FunctionComponent<IDemoProps> = (props) => {
	// State hooks
	const [lists, setLists] = useState<IListInfo[]>([])
	const [selectedListItems, setSelectedListItems] = useState<IListItem[]>([])
	const [selectedListTitle, setSelectedListTitle] = useState<string>('')
	const [userInfo, setUserInfo] = useState<IUserInfo | undefined>(undefined)
	const [siteInfo, setSiteInfo] = useState<ISiteInfo | undefined>(undefined)

	// Loading states
	const [listsLoading, setListsLoading] = useState<boolean>(false)
	const [itemsLoading, setItemsLoading] = useState<boolean>(false)
	const [userInfoLoading, setUserInfoLoading] = useState<boolean>(false)

	// Error states
	const [listsError, setListsError] = useState<string | undefined>(undefined)
	const [itemsError, setItemsError] = useState<string | undefined>(undefined)
	const [userInfoError, setUserInfoError] = useState<string | undefined>(undefined)

	// UI state
	const [currentView, setCurrentView] = useState<CurrentView>('welcome')

	// Memoized SharePoint service
	const sharePointService = useMemo(
		() => new SharePointService(props.context),
		[props.context]
	)

	// Load user and site information
	const loadUserAndSiteInfo = useCallback(async (): Promise<void> => {
		setUserInfoLoading(true)
		setUserInfoError(undefined)

		try {
			const [userInfoResult, siteInfoResult] = await Promise.all([
				sharePointService.getCurrentUser(),
				sharePointService.getSiteInfo(),
			])

			setUserInfo(userInfoResult)
			setSiteInfo(siteInfoResult)
			setUserInfoLoading(false)
		} catch (error) {
			setUserInfoError(
				(error as Error).message || 'Failed to load user and site information'
			)
			setUserInfoLoading(false)
		}
	}, [sharePointService])

	// Load SharePoint lists
	const loadLists = useCallback(async (): Promise<void> => {
		setListsLoading(true)
		setListsError(undefined)

		try {
			const listsResult = await sharePointService.getLists()
			setLists(listsResult)
			setListsLoading(false)
			setCurrentView('lists')
		} catch (error) {
			setListsError((error as Error).message || 'Failed to load lists')
			setListsLoading(false)
		}
	}, [sharePointService])

	// Load list items
	const loadListItems = useCallback(
		async (listTitle: string): Promise<void> => {
			setItemsLoading(true)
			setItemsError(undefined)
			setSelectedListTitle(listTitle)

			try {
				const items = await sharePointService.getListItems(listTitle, 50)
				setSelectedListItems(items)
				setItemsLoading(false)
				setCurrentView('items')
			} catch (error) {
				setItemsError((error as Error).message || 'Failed to load list items')
				setItemsLoading(false)
			}
		},
		[sharePointService]
	)

	// Create new list item
	const createListItem = useCallback(
		async (itemData: { Title: string; Description?: string }): Promise<void> => {
			try {
				const success = await sharePointService.createListItem(
					selectedListTitle,
					itemData
				)

				if (success) {
					// Reload items to show the new item
					await loadListItems(selectedListTitle)
				} else {
					throw new Error('Failed to create item')
				}
			} catch (error) {
				console.error('Error creating item:', error)
				throw error
			}
		},
		[sharePointService, selectedListTitle, loadListItems]
	)

	// Back to lists navigation
	const onBackToLists = useCallback((): void => {
		setCurrentView('lists')
		setSelectedListItems([])
		setSelectedListTitle('')
		setItemsError(undefined)
	}, [])

	// Load user info on component mount
	useEffect(() => {
		loadUserAndSiteInfo().catch(console.error)
	}, [loadUserAndSiteInfo])

	// Welcome view render function
	const renderWelcomeView = useCallback((): JSX.Element => {
		const {
			description,
			isDarkTheme,
			environmentMessage,
			hasTeamsContext,
			userDisplayName,
		} = props

		return (
			<section className={`${styles.demo} ${hasTeamsContext ? styles.teams : ''}`}>
				<div className={styles.welcome}>
					<img
						alt=''
						src={
							isDarkTheme
								? require('../assets/welcome-dark.png')
								: require('../assets/welcome-light.png')
						}
						className={styles.welcomeImage}
					/>
					<h2>Welcome, {escape(userDisplayName)}!</h2>
					<div>{environmentMessage}</div>
					<div>
						Web part property: <strong>{escape(description)}</strong>
					</div>
				</div>

				<Stack tokens={{ childrenGap: 20 }}>
					<div>
						<h3>SharePoint Framework Demo Features</h3>
						<p>
							This comprehensive SPFx demo showcases various SharePoint integration
							capabilities:
						</p>

						<ul className={styles.links}>
							<li>
								<strong>Lists Management:</strong> View and interact with SharePoint
								lists
							</li>
							<li>
								<strong>CRUD Operations:</strong> Create, read, update, and delete
								list items
							</li>
							<li>
								<strong>User Profile:</strong> Display current user and site
								information
							</li>
							<li>
								<strong>Modern UI:</strong> Built with Fluent UI React components
							</li>
							<li>
								<strong>Responsive Design:</strong> Works on desktop, tablet, and
								mobile devices
							</li>
						</ul>
					</div>

					<div>
						<h4>Learn more about SPFx development:</h4>
						<ul className={styles.links}>
							<li>
								<a href='https://aka.ms/spfx' target='_blank' rel='noreferrer'>
									SharePoint Framework Overview
								</a>
							</li>
							<li>
								<a
									href='https://aka.ms/spfx-yeoman-graph'
									target='_blank'
									rel='noreferrer'
								>
									Use Microsoft Graph in your solution
								</a>
							</li>
							<li>
								<a
									href='https://aka.ms/spfx-yeoman-teams'
									target='_blank'
									rel='noreferrer'
								>
									Build for Microsoft Teams using SharePoint Framework
								</a>
							</li>
							<li>
								<a
									href='https://aka.ms/spfx-yeoman-viva'
									target='_blank'
									rel='noreferrer'
								>
									Build for Microsoft Viva Connections using SharePoint Framework
								</a>
							</li>
						</ul>
					</div>
				</Stack>
			</section>
		)
	}, [props])

	// Current view render function
	const renderCurrentView = useCallback((): JSX.Element => {
		switch (currentView) {
			case 'lists':
				return (
					<ListViewer
						lists={lists}
						loading={listsLoading}
						error={listsError}
						onRefresh={loadLists}
						onListSelect={loadListItems}
					/>
				)

			case 'items':
				return (
					<ItemViewer
						listTitle={selectedListTitle}
						items={selectedListItems}
						loading={itemsLoading}
						error={itemsError}
						onRefresh={() => loadListItems(selectedListTitle)}
						onBack={onBackToLists}
						onCreateItem={createListItem}
					/>
				)

			case 'profile':
				return (
					<UserInfo
						userInfo={userInfo}
						siteInfo={siteInfo}
						loading={userInfoLoading}
						error={userInfoError}
					/>
				)

			default:
				return renderWelcomeView()
		}
	}, [
		currentView,
		lists,
		selectedListItems,
		selectedListTitle,
		listsLoading,
		itemsLoading,
		listsError,
		itemsError,
		userInfo,
		siteInfo,
		userInfoLoading,
		userInfoError,
		loadLists,
		loadListItems,
		onBackToLists,
		createListItem,
		renderWelcomeView,
	])

	// Handle pivot navigation
	const handlePivotLinkClick = useCallback(
		(item?: PivotItem): void => {
			const key = item?.props.itemKey as CurrentView
			if (key === 'lists' && lists.length === 0) {
				loadLists().catch((error) => {
					console.error('Error loading lists:', error)
				})
			} else {
				setCurrentView(key)
			}
		},
		[lists.length, loadLists]
	)

	// Main render
	return (
		<div className={styles.demo}>
			<Stack tokens={{ childrenGap: 20 }}>
				<Text variant='xxLarge' style={{ fontWeight: 'bold', textAlign: 'center' }}>
					SharePoint Framework Demo
				</Text>

				<Pivot selectedKey={currentView} onLinkClick={handlePivotLinkClick}>
					<PivotItem headerText='Welcome' itemKey='welcome' />
					<PivotItem headerText='Lists' itemKey='lists' />
					<PivotItem headerText='User Profile' itemKey='profile' />
				</Pivot>

				{renderCurrentView()}
			</Stack>
		</div>
	)
}

export default Demo

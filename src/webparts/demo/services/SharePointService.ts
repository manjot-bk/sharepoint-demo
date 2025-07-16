import { WebPartContext } from '@microsoft/sp-webpart-base'
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http'
import { IListItem } from '../models/IListItem'

export interface IListInfo {
	Title: string
	Id: string
	ItemCount: number
	Created: string
	BaseType: number
}

export interface ISearchResult {
	Title: string
	Path: string
	Description: string
	Author: string
	LastModifiedTime: string
}

export interface IUserInfo {
	Id: number
	Title: string
	Email: string
	LoginName: string
}

export interface ISiteInfo {
	Title: string
	Description: string
	Url: string
	Created: string
}

export class SharePointService {
	private _context: WebPartContext

	constructor(context: WebPartContext) {
		this._context = context
	}

	// Get all lists in the current site
	public async getLists(): Promise<IListInfo[]> {
		try {
			const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title,Id,ItemCount,Created,BaseType&$filter=Hidden eq false`

			const response: SPHttpClientResponse = await this._context.spHttpClient.get(
				endpoint,
				SPHttpClient.configurations.v1
			)

			if (response.ok) {
				const data = await response.json()
				return data.value
			}
			return []
		} catch (error) {
			console.error('Error getting lists:', error)
			return []
		}
	}

	// Get list items with pagination
	public async getListItems(
		listTitle: string,
		top: number = 10
	): Promise<IListItem[]> {
		try {
			const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,Title,Created,Modified,Author/Title,Author/Email&$expand=Author&$top=${top}`

			const response: SPHttpClientResponse = await this._context.spHttpClient.get(
				endpoint,
				SPHttpClient.configurations.v1
			)

			if (response.ok) {
				const data = await response.json()
				return data.value.map(
					(item: {
						Id: number
						Title: string
						Created: string
						Modified: string
						Author?: { Title: string; Email: string }
					}) => ({
						Id: item.Id,
						Title: item.Title,
						Created: item.Created,
						Modified: item.Modified,
						Author: {
							Title: item.Author?.Title || 'Unknown',
							Email: item.Author?.Email || '',
						},
					})
				)
			}
			return []
		} catch (error) {
			console.error('Error getting list items:', error)
			return []
		}
	}

	// Create a new list item
	public async createListItem(
		listTitle: string,
		itemData: Record<string, string | number>
	): Promise<boolean> {
		try {
			const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items`

			const response: SPHttpClientResponse = await this._context.spHttpClient.post(
				endpoint,
				SPHttpClient.configurations.v1,
				{
					headers: {
						Accept: 'application/json;odata=nometadata',
						'Content-type': 'application/json;odata=nometadata',
						'odata-version': '',
					},
					body: JSON.stringify(itemData),
				}
			)

			return response.ok
		} catch (error) {
			console.error('Error creating list item:', error)
			return false
		}
	}

	// Get current user information
	public async getCurrentUser(): Promise<IUserInfo | undefined> {
		try {
			const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web/currentuser`

			const response: SPHttpClientResponse = await this._context.spHttpClient.get(
				endpoint,
				SPHttpClient.configurations.v1
			)

			if (response.ok) {
				const user = await response.json()
				return {
					Id: user.Id,
					Title: user.Title,
					Email: user.Email,
					LoginName: user.LoginName,
				}
			}
			return undefined
		} catch (error) {
			console.error('Error getting current user:', error)
			return undefined
		}
	}

	// Get site information
	public async getSiteInfo(): Promise<ISiteInfo | undefined> {
		try {
			const endpoint = `${this._context.pageContext.web.absoluteUrl}/_api/web?$select=Title,Description,Url,Created`

			const response: SPHttpClientResponse = await this._context.spHttpClient.get(
				endpoint,
				SPHttpClient.configurations.v1
			)

			if (response.ok) {
				const site = await response.json()
				return {
					Title: site.Title,
					Description: site.Description,
					Url: site.Url,
					Created: site.Created,
				}
			}
			return undefined
		} catch (error) {
			console.error('Error getting site info:', error)
			return undefined
		}
	}
}

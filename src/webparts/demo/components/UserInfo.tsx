import * as React from 'react'
import {
	Persona,
	PersonaSize,
	Stack,
	Text,
	Spinner,
	SpinnerSize,
	MessageBar,
	MessageBarType,
} from '@fluentui/react'
import { IUserInfo, ISiteInfo } from '../services/SharePointService'
import styles from './Demo.module.scss'

export interface IUserInfoProps {
	userInfo?: IUserInfo
	siteInfo?: ISiteInfo
	loading: boolean
	error?: string
}

export const UserInfo: React.FunctionComponent<IUserInfoProps> = (props) => {
	const { userInfo, siteInfo, loading, error } = props

	if (loading) {
		return (
			<div className={styles.center}>
				<Spinner
					size={SpinnerSize.medium}
					label='Loading user and site information...'
				/>
			</div>
		)
	}

	if (error) {
		return (
			<MessageBar messageBarType={MessageBarType.error}>
				Error loading information: {error}
			</MessageBar>
		)
	}

	return (
		<div>
			{userInfo && (
				<div className={styles.userProfile}>
					<Stack tokens={{ childrenGap: 15 }}>
						<Text variant='large' style={{ fontWeight: 'bold' }}>
							Current User
						</Text>

						<Persona
							text={userInfo.Title}
							secondaryText={userInfo.Email}
							size={PersonaSize.size72}
							imageAlt={userInfo.Title}
						/>

						<Stack tokens={{ childrenGap: 5 }}>
							<Text>
								<strong>Display Name:</strong> {userInfo.Title}
							</Text>
							<Text>
								<strong>Email:</strong> {userInfo.Email}
							</Text>
							<Text>
								<strong>Login Name:</strong> {userInfo.LoginName}
							</Text>
							<Text>
								<strong>User ID:</strong> {userInfo.Id}
							</Text>
						</Stack>
					</Stack>
				</div>
			)}

			{siteInfo && (
				<div className={styles.siteInfo}>
					<Stack tokens={{ childrenGap: 15 }}>
						<Text variant='large' style={{ fontWeight: 'bold' }}>
							Site Information
						</Text>

						<Stack tokens={{ childrenGap: 5 }}>
							<Text>
								<strong>Site Title:</strong> {siteInfo.Title}
							</Text>
							<Text>
								<strong>Description:</strong>{' '}
								{siteInfo.Description || 'No description available'}
							</Text>
							<Text>
								<strong>URL:</strong> {siteInfo.Url}
							</Text>
							<Text>
								<strong>Created:</strong>{' '}
								{new Date(siteInfo.Created).toLocaleString()}
							</Text>
						</Stack>
					</Stack>
				</div>
			)}
		</div>
	)
}

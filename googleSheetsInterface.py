from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage
import argparse

import httplib2

SCOPES = "https://www.googleapis.com/auth/drive.file"
CLIENT_SECRET_FILE = "client_secret.json"
APPLICATION_NAME = "SWGOH Miner"
CREDENTIAL_PATH = "googleSheetsCredentials.json"

FOLDER_TYPE = "application/vnd.google-apps.folder"
SHEETS_TYPE = "application/vnd.google-apps.spreadsheet"


class SpreadsheetsInterface:
	global FOLDER_TYPE
	global SHEETS_TYPE

	def __init__(self, guildName):
		self.credentials = self.getCredentials()
		self.driveService = self.getService("drive")
		self.spreadsheetsService = self.getService("sheets")
		self.guildFolderID = self.findFile(guildName, FOLDER_TYPE)
		if self.guildFolderID is None:
			self.guildFolderID = self.driveService.files().create(body={
				"name": guildName,
				"mimeType": FOLDER_TYPE
			}, fields="id").execute().get("id")

	def getCredentials(self):
		store = Storage(CREDENTIAL_PATH)
		credentials = store.get()
		if not credentials or credentials.invalid:
			flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
			flow.user_agent = APPLICATION_NAME
			flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
			credentials = tools.run_flow(flow, store, flags)
		return credentials

	def getService(self, type):
		http = self.credentials.authorize(httplib2.Http())
		if type == "drive":
			return discovery.build("drive", "v3", http=http)
		elif type == "sheets":
			discoveryURL = ("https://sheets.googleapis.com/$discovery/rest?version=v4")
			return discovery.build("sheets", "v4", http=http, discoveryServiceUrl=discoveryURL)
		else:
			raise ValueError("Didn't specify a valid service type")

	def findFile(self, fileName, mimeType, extra=""):
		pageToken = None
		while True:
			response = self.driveService.files().list(
				q="name='{}' and mimeType='{}' and trashed=false".format(fileName, mimeType) + extra,
				spaces="drive", fields="nextPageToken, files(id, name)", pageToken=pageToken).execute()
			if len(response.get("files", [])) == 1:
				return response.get("files", [])[0].get("id")
			elif len(response.get("files", [])) > 1:
				raise ValueError("More than 1 instance of file")
			pageToken = response.get("nextPageToken", None)
			if pageToken is None:
				break
		return None

	def writeData(self, playerName, rows, sheetName, rowCount):
		# Create player spreadhsheet if it doesn't exist
		playerSpreadsheetID = self.findFile(playerName, SHEETS_TYPE, " and '{}' in parents".format(self.guildFolderID))
		if playerSpreadsheetID is None:
			playerSpreadsheetID = self.driveService.files().create(body={
				"name": playerName,
				"mimeType": SHEETS_TYPE,
				"parents": [self.guildFolderID]
			}, fields="id").execute().get("id")
		# Create data sheet if it doesn't exists
		playerSheets = self.spreadsheetsService.spreadsheets().get(spreadsheetId=playerSpreadsheetID).execute().get("sheets")
		sheetID = None
		startingSheetExists = False
		for sheet in playerSheets:
			if sheet["properties"]["title"] == sheetName:
				sheetID = sheet["properties"]["sheetId"]
			elif sheet["properties"]["sheetId"] == 0:
				startingSheetExists = True
		if sheetID is None:
			response = self.spreadsheetsService.spreadsheets().batchUpdate(spreadsheetId=playerSpreadsheetID, body={
				"requests": [{
					"addSheet": {
						"properties": {
							"title": sheetName,
							"gridProperties": {
								"rowCount": rowCount,
								"columnCount": 2
							}
						}
					}
				}]
			}).execute()
			sheetID = response.get("replies")[0]["addSheet"]["properties"]["sheetId"]
			self.spreadsheetsService.spreadsheets().batchUpdate(spreadsheetId=playerSpreadsheetID, body={
				"requests": [{
					"addNamedRange": {
						"namedRange": {
							"name": sheetName.replace(" ", "_"),
							"range": {
								"sheetId": sheetID
							}
						}
					}
				}, {
					"addProtectedRange": {
						"protectedRange": {
							"range": {
								"sheetId": sheetID
							},
							"warningOnly": True
						}
					}
				}]
			}).execute()
		if startingSheetExists:
			self.spreadsheetsService.spreadsheets().batchUpdate(spreadsheetId=playerSpreadsheetID, body={
				"requests": [{
					"deleteSheet": {
						"sheetId": 0
					}
				}]
			}).execute()
		# Write the data to the sheet
		self.spreadsheetsService.spreadsheets().batchUpdate(spreadsheetId=playerSpreadsheetID, body={
			"requests": [{
				"updateSheetProperties": {
					"properties": {
						"sheetId": sheetID,
						"gridProperties": {
							"rowCount": rowCount,
							"columnCount": 2
						}
					},
					"fields": "gridProperties"
				}
			}, {
				"updateCells": {
					"start": {
						"sheetId": sheetID,
						"rowIndex": 0,
						"columnIndex": 0
					},
					"rows": rows,
					"fields": "userEnteredValue,userEnteredFormat.textFormat"
				}
			}]
		}).execute()

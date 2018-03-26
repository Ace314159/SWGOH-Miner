import urllib.request
import urllib.error

from googleSheetsInterface import SpreadsheetsInterface
from bs4 import BeautifulSoup
import pickle
import os

try:
	with open("playersDone.pkl", "rb") as f:
		playersDone = pickle.load(f)
except FileNotFoundError:
	open("playersDone.pkl", "w").close()
	playersDone = []


def finishedPlayer(playerName):
	playersDone.append(playerName)
	with open("playersDone.pkl", "wb") as f:
		pickle.dump(playersDone, f)


def getHtml(URL):
	request = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
	return urllib.request.urlopen(request)


def getFormat(val):
	try:
		int(val)
		return "numberValue"
	except ValueError:
		return "stringValue"


def getRowFromValues(val1="", val2="", bold1=False, bold2=False, ital1=False, ital2=False):
	format1 = getFormat(val1)
	format2 = getFormat(val2)
	return {"values": [
		{"userEnteredValue": {format1: val1}, "userEnteredFormat": {"textFormat": {"bold": bold1, "italic": ital1}}},
		{"userEnteredValue": {format2: val2}, "userEnteredFormat": {"textFormat": {"bold": bold2, "italic": ital2}}}
	]}


def writeToSheets(playerName, data, sheetName):
	rows = []
	rowCount = 0
	if type(data) is dict:
		for key, value in data.items():
			rows.append(getRowFromValues(val1=key, val2=value))
			rowCount += 1
	elif type(data) is list:
		first = True
		for item in data:
			if not first:
				rows.append(getRowFromValues())
				rowCount += 1
			rows.append(getRowFromValues(val1=item.pop("Name"), bold1=True))
			rowCount += 1
			for key, value in item.items():
				if type(value) is list:
					if(len(value) > 0):
						rows.append(getRowFromValues(val1=key, val2=value[0], ital1=True))
					else:
						rows.append(getRowFromValues(val1=key, val2="None"))
					rowCount += 1
					for i in value[1:]:
						rows.append(getRowFromValues(val1="", val2=i))
						rowCount += 1
				else:
					rows.append(getRowFromValues(val1=key, val2=value))
					rowCount += 1
			first = False

	else:
		raise ValueError("Type of data is not dict or list")
	spreadsheetsInterface.writeData(playerName, rows, sheetName, rowCount)


domain = "https://swgoh.gg"
URL = domain + "/g/14739/warrior-knights/"
guildName = "Warrior Knights"
charactersDir = "collection"
shipsDir = "ships"
modsDir = "mods"
modShapes = [None, "Transmitter", "Processor", "Data-Bus", "Receiver", "Holo-Array", "Multiplexer"]
modSets = [None, "Health", "Defense", "Crit Damage", "Crit Chance", "Tenacity", "Offense", "Potency", "Speed"]
print("Preparing the Google Sheets")
spreadsheetsInterface = SpreadsheetsInterface(guildName)

statsSheetName = "Stats"
playerInfoSheetName = "Player Info"
charactersSheetName = "Characters"
shipsSheetName = "Ships"
modsSheetName = "Mods"

print("Getting list of users in the guild")
# Get the html of the page
html = getHtml(URL)
soup = BeautifulSoup(html, "html.parser")

# Parse the html to get all the players
playerTable = soup.find("table")
for playerDir in playerTable.findAll("a"):
	# Main Profile Page
	playerURL = domain + playerDir.get("href")
	playerName = playerDir.get_text().strip()
	playerHtml = getHtml(playerURL)
	playerSoup = BeautifulSoup(playerHtml, "html.parser")
	if playerName in playersDone:
		print("Skipping", playerName)
		continue
	print()
	print("Starting", playerName)

	print("Getting Main Profile")
	# Stats
	statsSoup = playerSoup.find("ul", class_="panel-menu m-0")
	stats = {}
	for stat in statsSoup.findAll("li"):
		raw = stat.get_text().strip().rpartition("\n")
		stats[raw[0]] = raw[2]
	writeToSheets(playerName, stats, statsSheetName)
	# Player Info
	infoSoup = playerSoup.findAll("div", class_="panel-profile")[1].div
	playerInfo = {}
	for info in infoSoup.findAll("p"):
		raw = info.get_text().strip().rpartition(" ")
		playerInfo[raw[0]] = raw[2]
	writeToSheets(playerName, playerInfo, playerInfoSheetName)

	# Characters Page
	print("Getting Characters")
	charactersURL = playerURL + charactersDir
	charactersHtml = getHtml(charactersURL)
	charactersSoup = BeautifulSoup(charactersHtml, "html.parser").find("li", class_="media").div
	characters = []
	for characterSoup in charactersSoup.findAll("div", recursive=False):
		if "collection-char-missing" in characterSoup.div.get("class"):
			break
		characterInfo = {}
		characterDivs = characterSoup.div.findAll("div", recursive=False)
		characterInfo["Star Number"] = str(7 - len(characterDivs[0].a.findAll("div", class_="star-inactive")))
		characterInfo["Level"] = characterDivs[0].a.find("div", class_="char-portrait-full-level").get_text()
		characterInfo["Gear Level"] = characterDivs[0].a.find("div", class_="char-portrait-full-gear-level").get_text()
		characterInfo["Power"] = characterDivs[1].get("title").split(" ", 1)[1]
		characterInfo["Name"] = characterDivs[2].get_text().strip()
		characters.append(characterInfo)
	writeToSheets(playerName, characters, charactersSheetName)

	# Ships Page
	print("Getting Ships")
	shipsURL = playerURL + shipsDir
	shipsHtml = getHtml(shipsURL)
	shipsSoup = BeautifulSoup(shipsHtml, "html.parser").find("li", class_="media").div
	ships = []
	for shipSoup in shipsSoup.findAll("div", recursive=False):
		if "collection-ship-missing" in shipSoup.div.get("class"):
			break
		shipInfo = {}
		shipDivs = shipSoup.div.findAll("div", recursive=False)
		shipInfo["Star Number"] = str(7 - len(shipDivs[0].findAll("div", class_="ship-portrait-full-star-inactive")))
		shipInfo["Level"] = shipDivs[0].find("div", class_="ship-portrait-full-frame-level").get_text()
		shipInfo["Crew"] = list(map(lambda x: x.get("alt"), shipDivs[0].findAll("img", class_="char-portrait-full-img")))
		shipInfo["Name"] = shipDivs[1].get_text().strip()
		ships.append(shipInfo)
	writeToSheets(playerName, ships, shipsSheetName)

	# Mods Page
	print("Getting Mods")
	mods = []
	pageNum = 1
	try:
		while True:
			modsURL = playerURL + modsDir + "?page=" + str(pageNum)
			modsHtml = getHtml(modsURL)
			modsSoup = BeautifulSoup(modsHtml, "html.parser").find("li", class_="media").div
			for modSoup in modsSoup.findAll("div", recursive=False):
				modInfo = {}
				modDivs = modSoup.div.div.findAll("div", recursive=False)
				modInfo["Pip Number"] = str(len(modDivs[0].div.span.findAll("span")))
				modInfo["Level"] = modDivs[0].div.get_text().strip()
				modInfo["Character"] = modDivs[0].div.a.div.get("title")
				modImg = modDivs[0].div.find("img", class_="statmod-img")
				modInfo["Name"] = modImg.get("alt")
				modImgSrc = (modImg.get("src").split("_"))
				modInfo["Set"] = modSets[int(modImgSrc[-2])]
				modInfo["Shape"] = modShapes[int(modImgSrc[-1][0])]
				rawStats = list(filter(bool, modDivs[1].div.get_text().strip().split("\n")))
				modInfo["Primary Stat"] = rawStats[0]
				modInfo["Secondary Stats"] = rawStats[1:]
				mods.append(modInfo)
			print("Finished Mods Page", pageNum, end="\r")
			pageNum += 1
	except urllib.error.HTTPError as e:
		print()
		if e.code != 404:
			raise e
	writeToSheets(playerName, mods, modsSheetName)
	finishedPlayer(playerName)

os.remove("playersDone.pkl")

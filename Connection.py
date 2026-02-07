import urllib.parse
from bs4 import BeautifulSoup
from requests_html import HTMLSession
from typing import List
from _datetime import datetime
from openpyxl import Workbook
from PIL import Image, ImageDraw, ImageFont
import os

_ROOT_URL = "https://dividendhistory.org"

current_date = {
	"year": datetime.today().year,
	"month": datetime.today().month,
	"day": datetime.today().day
}
full_date = str(current_date["year"]) + "-" + str(current_date["month"]) + "-" + str(current_date["day"])

_options = {
	"monthly": r'https://dividendhistory.org/monthly-payout/',
	"weekly": r'https://dividendhistory.org/weekly-payout/'
}
_country_options = {
	"USA": "us",
	"CAN": "ca",
	"ANY": ".."
}

# Start tallying points (Personal scores, change as needed)
# - Consistency (No increase or decrease)
# - Increases (Save a tally of increases)
# - Decreases (Save a tally of decreases)
# - Change_Over_Time (Increase_Tally + Decrease_Tally)
# - Change_Over_Year (Increase/Decrease within this year)
points = {
	"dividend_changes_over_time":{
		"Consistency": 3,
		"Increase": 2,
		"Decrease": -36,
	},
	"dividend_changes_over_year":{
		"Consistency": 4,
		"Increase": 4,
		"Decrease": -60,
	},
	"change_over_time":{
		"large_tally_positive": 2,
		"tally_positive_neutral": 1,
		"tally_negative": -3,
		"large_tally_negative": -6
	},
	"change_over_year":{
		"large_tally_positive": 3,
		"tally_positive_neutral": 2,
		"tally_negative": -25,
		"large_tally_negative": -900
	}
}
change_thresholds = {
	"change_over_time_thresholds":{
		"greater_than_large_positive": 50,
		"greater_than_positive_neutral": 0,
		"less_than_negative": 0,
		"less_than_large_negative": -50,
	},
	"change_over_year_thresholds":{
		"greater_than_large_positive": 40,
		"greater_than_positive_neutral": 0,
		"less_than_negative": 0,
		"less_than_large_negative": -10
	}
}

inp = _options["monthly"]

class Stock:
	def __init__(self, link, code, name, country, yield_, ex_div_date):
		self.link = link
		self.code = code
		self.name = name
		self.country = country
		self.yield_ = yield_
		self.ex_div_date = ex_div_date
		self.price = -1.11
		self.per_stock_div = -1.11
		self.price_div_percent = -1.11 # Lower is better, if percent is lower, div is giving a larger portion of returns
		self.points = 0 # Neutral, >0 Add to list, <0 Remove from list

def div_site_soup(url) -> BeautifulSoup:
	'''
	Connect to the website, and create the soup so we can sift for the stock information.
	Return the created BeautifulSoup.
	'''

	# Pretending like the code is a person as to not be blocked.
	header = {
        'Accept': 'image/avif,image/webp,image/png,image/svg+xml,image/*;q=0.8,*/*;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br, zstd',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:141.0) Gecko/20100101 Firefox/141.0',
        'Host': urllib.parse.urlparse(url).netloc,
        'Accept-Language': 'en-US,en;q=0.5',
        'Referer': _ROOT_URL,
        'Connection': 'keep-alive'
    }

	# Load the page, and render the page to load all information for buttons that do queries.
	session = HTMLSession()
	request = session.get(url, stream=True, timeout=None, headers=header)
	try:
		request.html.render(timeout=60)
	except Exception:
		pass

	# Save the soup close the connection so the pc doesn't lag again...
	soup = BeautifulSoup(request.content, "html.parser")
	session.close()

	# Return the formatted soup of the render.
	return soup

def setup_table_list(soup, chosen_country) -> List[Stock]:
	'''
	Takes all possible stocks and format them from the html into stock objects.
	We'll use these to gather their prices and history.
	Returns a list of all stock objects.
	'''

	formatted_stock_list = []
	stock_list_raw = soup.find_all("tbody")[0].find_all("tr")
	for stock in stock_list_raw:
		# Splits a stock into
		# 0 - href and code
		# 1 - name and country
		# 2 - yield percentage
		# 3 - next ex-div date
		elements = stock.find_all("td")

		# Can be either 'us' or 'ca' | For filters on information
		country = elements[0].find(class_=True)['class'][2].rsplit('-')[2]

		# Early check if we want to add the stock to the list based on user input
		if _country_options[chosen_country] == ".." or country == _country_options[chosen_country]:
			# Link - website url + href isolated | To get price and history
			link = _ROOT_URL + elements[0].find(href=True)['href']

			# Code in 0 element | Formatting the table
			code = elements[0].a.contents[0]

			# Company name | Formatting the table
			name = elements[1].contents[0]

			# Yield of the stock | Formatting the table
			yield_ = elements[2].contents[0]

			# The next ex div date | Important information for someone investing
			try:
				ex_div_date = elements[3].contents[0]
			except Exception:
				ex_div_date = "0000-00-00"

			# Formalize stock and add to formatted list
			stock_obj = Stock(link, code, name, country, yield_, ex_div_date)
			formatted_stock_list.append(stock_obj)
	return formatted_stock_list

def within_year(date) -> str:
	'''
	Check entry is within the year, apply bonuses in dictionary if it is.
	'''
	if (date[0] == current_date["year"] - 1 and date[1] >= current_date["month"]) or\
			(date[0] == current_date["year"] and date[1] <= current_date["month"]):
		return "year"
	return "time"

def change_points(time_interval, tally) -> int:
	'''
	Selects the time interval in the above dictionaries to assign points proper to the indicated interval.
	Compares the tally to the dictionary settings and returns the appropriate point value as an int.
	'''
	# Chosen time interval set in option
	option = "change_over_" + time_interval + "_thresholds"
	interval = "change_over_" + time_interval

	# Return points based on tally value (Order matters) (Higher or lower than smallest threshold, then checks larger)
	if change_thresholds[option]["greater_than_positive_neutral"] <= tally:
		if change_thresholds[option]["greater_than_large_positive"] < tally:
			return points[interval]["large_tally_positive"]
		return points[interval]["tally_positive_neutral"]
	else:
		if change_thresholds[option]["less_than_large_negative"] > tally:
			return points[interval]["large_tally_negative"]
		return points[interval]["tally_negative"]

def stock_points(history_soup_list) -> int:
	'''
	Takes in the history of a stock and scores it based on the following:
		- How many times the dividend payout changes.
		- The overall percentage change.
		(Biased for entries within one year of current time)
	'''
	# Initial variables
	change_over_year_lock = True
	stock_points = 0
	tally = 0

	# Go over every entry in the history
	for entry in history_soup_list:
		# Bonus if within the year (Usually more important)
		entry_date = (int(entry.find("td").contents[0][0:4]), int(entry.find("td").contents[0][5:7]))
		change_option = within_year(entry_date)

		# Once out of the year, do a tally check to give change over year points
		if change_option == "time" and change_over_year_lock:
			stock_points += change_points("year", tally)
			change_over_year_lock = False

		# Change is for identifying changes, div_change_opt is to shorten the dictionary call
		change = entry.find("span", class_=True)
		div_change_opt = "dividend_changes_over_" + change_option

		# Either goes up (change.attrs), down (Else), or stays neutral (None)
		if change is None:
			stock_points += points[div_change_opt]["Consistency"]
		elif change.attrs['class'][1] == "percent-increase":
			stock_points += points[div_change_opt]["Increase"]
			tally += float(entry.find("span", class_=True).contents[1].replace("%", ""))
		else:
			stock_points += points[div_change_opt]["Decrease"]
			tally += float(entry.find("span", class_=True).contents[1].replace("%", ""))

	# All entries have finished, change over time points
	stock_points += change_points("time", tally)
	return stock_points

def stock_profile(stock) -> None:
	'''
	Looks at stock's dedicated page and gathers:
		- The last close price.
		- The most recent dividend payout.
		- The percentage that the most recent dividend payout is compared to the last close price.
		- The score of the stock, calculated using values assigned in 'points' by user's importance of those categories.
			- Notes:
				- change_thresholds are % changes in dividend payouts over time, change by preference.
				- Scores calculated < 0 will be deleted from the list and will not be in the final table.
	'''
	# Look at page
	profile = div_site_soup(stock.link)

	# Grab Price
	stock.price = float(str(profile.find(lambda tag:tag.name=="p" and "Last Close Price" in tag.text).contents[0]).rsplit('$')[1])

	# Get history, but nothing unconfirmed
	history = profile.find_all("tbody")[-1].find_all(lambda tag: tag.name == "tr" and 'class' in tag.attrs and "unconfirmed-div" not in tag['class'])

	# Most recent payout
	stock.per_stock_div = float(str(history[0].find_all("td")[2].contents[0][1:]).rsplit(" ")[0])

	# A percent that indicates larger portion of returns the lower the number
	stock.price_div_percent = stock.price/stock.per_stock_div

	# Tallying points
	stock.points = stock_points(history)
	print("Stock Profile: " + stock.name + " Complete")
	return

def all_stocks_profile(stock_list) -> None:
	'''
	Iterates over the stock list and runs it through the profile function. Sorts them by price_div_percent.
	(How many stocks needed before you can buy another stock with the payout)
	'''
	for stock in stock_list:
		stock_profile(stock)
	stock_list.sort(key=lambda x: x.price_div_percent)

def clean_scores(stock_list) -> List[Stock]:
	return [stock for stock in stock_list if stock.points > 0 and stock.price_div_percent < 125]

def make_dirs(full_date) -> str:
	'''
	Makes directories if they don't already exist to hold specifically today's information.
	'''
	os.makedirs(os.getcwd() + "\\DivScores", exist_ok=True)
	os.makedirs(os.getcwd() + "\\DivScores\\" + full_date, exist_ok=True)
	return os.getcwd() + "\\DivScores\\" + full_date

def to_spreadsheet(stock_list) -> None:
	'''
	Takes the list of information and creates a readable spreadsheet.
	'''
	# Setup a spreadsheet
	workbook = Workbook()
	sheet = workbook.active
	sheet.title = "Dividend Scorings For " + full_date
	sheet.append(["Rank", "Payout to Stock Price", "Score", "Code", "Name", "Country",
				  "Price", "Per Stock Payout", "Yield", "Next Ex-Div Date", "URL"])

	# Make appropriate directories and get the directory to save the excel in
	save_location = make_dirs(full_date)

	# Add stock information
	rank = 1
	for stock in stock_list:
		if stock.points > 0:
			sheet.append([rank, stock.price_div_percent, stock.points, stock.code, stock.name, stock.country,
						  stock.price, stock.per_stock_div, stock.yield_, stock.ex_div_date, stock.link])
			rank += 1

	workbook.save(save_location + "\\Div-Score-" + full_date + ".xlsx")
	return

def longest_width(stock_list, category, title, font) -> int:
	'''
	Gets the longest width of the column so that the image drawer can calculate spacing.
	'''
	#temp image to get lengths from
	tmp_img = Image.new('RGB', (300, 300), color="white")
	draw = ImageDraw.Draw(tmp_img)

	# Title length in case it's longer than anything in stocks or number of stocks
	title_width = draw.textlength(str(title), font=font)
	if category is None:
		# Case for rankings, return early so there isn't an error
		width = draw.textlength(str(len(stock_list)), font=font)
		return int(width)+8 if width > title_width else int(title_width)+8
	else:
		# Grabs the stock with the longest character-wise attribute of the given category
		width = max(stock_list, key=lambda stock: draw.textlength(str(getattr(stock, category)), font=font))

	# Accounts for the title being longer than the longest attribute
	if draw.textlength(str(getattr(width, category)), font=font) < title_width:
		width = int(title_width)
	else:
		width = int(draw.textlength(str(getattr(width, category)), font=font))
	tmp_img.close()
	return width + 8

def to_image(stock_list) -> None:
	'''
	Draw the spreadsheet for viewing. Rows are the number of entries. Columns are always static (11).
	'''
	# Column titles
	columns = {
		"Rank": None,
		"Payout to Stock Price": "price_div_percent",
		"Score": "points",
		"Code":"code",
		"Name":"name",
		"Country":"country",
		"Price":"price",
		"Per Stock Payout":"per_stock_div",
		"Yield":"yield_",
		"Next Ex-Div Date":"ex_div_date",
		"URL":"link"
	}

	# Calculations used for the image
	font_size = 12
	font = ImageFont.load_default(font_size)

	cell_widths = []
	for category,title in zip(columns.values(),columns.keys()):
		cell_widths.append(longest_width(stock_list, category, title, font))

	cell_height = font_size + 8
	rows = len(stock_list) + 1

	# Setting up image dimensions
	image_width = sum(cell_widths)
	image_height = rows * cell_height

	# Creating an image to host
	img = Image.new('RGB', (image_width,image_height), color="white")
	draw = ImageDraw.Draw(img)

	#Drawing grid
	for horizontal in range(rows+ 1):
		height = horizontal * cell_height
		draw.line((0, height, image_width, height), fill="black")
	for vertical in range(len(cell_widths) + 1):
		if vertical == 0:
			draw.line((0, 0, 0, image_height), fill="black")
		else:
			width = sum(cell_widths[0:vertical])
			draw.line((width, 0, width, image_height), fill="black")

	#Draw column titles
	y = (cell_height - font_size) / 2
	x = 2
	for column_width, info in zip(cell_widths, columns.keys()):
		draw.text((x,y), info, fill="black", font=font)
		x += column_width
	y += cell_height

	#Drawing stocks
	rankings = 1
	for stock in stock_list:
		x = 2
		for column_width,info in zip(cell_widths,columns.values()):
			if info is None:
				text = str(rankings)
				rankings += 1
			else:
				text = str(getattr(stock, info))
			draw.text((x, y), text, fill="black", font=font)
			x += column_width
		y += cell_height

	# Make appropriate directories and get the directory to save the img in, then save it
	save_location = make_dirs(full_date)
	img.save(save_location + "\\Div-Score-" + full_date + ".png")

def testing() -> None:
	return

if __name__ == "__main__":
	country_choice = input("Choose a country (usa, can, any): ")
	if country_choice is None or country_choice.upper() not in _country_options.keys():
		country_choice = "ANY"
	else:
		country_choice = country_choice.upper()

	# Setup initial tables and stock objects with their scorings
	table = setup_table_list(div_site_soup(inp), country_choice)
	all_stocks_profile(table)

	# Get rid of negative scores and create a spreadsheet and image
	table = clean_scores(table)
	to_spreadsheet(table)
	to_image(table)

#CharmedTracker_V3.py
# Default Packages
import json
import pathlib
import traceback
from datetime import datetime
from datetime import timedelta
import json
import requests
import math
import logging
import os
import os.path
import csv
import pickle

# Packages from PiP
import openpyxl as pyxl
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

class Loggable:
	logger = logging.getLogger()

# Start:
# ct = CharmedTracker().main()
class CharmedTracker(Loggable):
	def __init__(self):
		self.config = Storage(filepath="./resources/config.json")
		self.orders_storage = StoredList(filepath="./resources/orders_storage.json")
		self.scans_storage = Storage(filepath="./resources/scans_storage.json", default_value=[])
		self.wms_api = WMS_API(config=self.config)
		self.google_api = GoogleSheets_API(config=self.config)

	def _remove_all_before_date(self, date: str):
		self.logger.info(f"Removing orders created before {date}")
		out = []
		for order in self.orders_storage.data:
			if not order.creation_date[:10] < date:
				out.append(order)
		self.orders_storage.data = out
		self.orders_storage.save()

	def _remove_all_after_date(self, date: str):
		self.logger.info(f"Removing orders created after {date}")
		out = []
		for order in self.orders_storage.data:
			if not order.creation_date[:10] > date:
				out.append(order)
		self.config.data["last_run_date"] = date + " 23:59:00" #tmp
		self.config.save()
		self.orders_storage.data = out
		self.orders_storage.save()

	def _set_all_to_unshipped(self):
		self.logger.info(f"Setting all orders to unshipped")
		for order in self.orders_storage.data:
			order.ship_status = None
			order.ship_date = None
		self.orders_storage.save()

	def main(self):
		#self.update_current_orders() #TODO remove
		new_orders = self.fetch_new_orders()
		matches_found = self.process_scans_folder()
		if new_orders or matches_found or True:#TODO remove
			self.logger.info("New orders/matches found. Updating Google Sheet...")
			self.update_google_sheet()
		else:
			self.logger.info("No new orders or matches found. Exiting")
		self.logger.info("CharmedTracker exit")

	def update_current_orders(self):
		match_flag = False
		time_limit = datetime.strftime(datetime.now() - timedelta(days=3), "%Y-%m-%d %H:%M:%S")
		for order in self.orders_storage.data:
			if order.close_date == None and order.creation_date < time_limit:
				match_flag = True
				order = self.wms_api.get_order(str(order.order_id))
		if match_flag:
			self.orders_storage.save()

	def fetch_new_orders(self):
		start_date = self.config.data["last_run_date"]
		end_date = now()
		new_orders_flag = False
		for customer in self.config.data["supported_customers"]:
			customer_id = str(self.config.data["supported_customers"][customer]["3PLC_customer_id"])
			orders_list = self.wms_api.get_3PLC_orders_since_date(customer_id, start_date, end_date)
			if orders_list:
				orders_count = str(len(orders_list))
				self.logger.info(f"{orders_count} results parsed")
				#
				for order in orders_list:
					'''
					if "letter" in order.tracking_number.lower():
						order = self.set_to_shipped(order)
						self.logger.info(f"Order {order.order_id} {order.tracking_number} set to shipped")
					'''
					if "cancel" not in order.reference_id.lower():
						self.orders_storage.add(order, index=order.creation_date)
				self.orders_storage.save()
				new_orders_flag = True
		if new_orders_flag:
			self.config.data["last_run_date"] = now()
			self.config.save()
			return new_orders_flag
	
	def process_scans_folder(self):
		parent_dir = pathlib.Path(os.getcwd()).parent.absolute()
		scans_dir = os.path.join(parent_dir, "scans")
		old_scans_dir = os.path.join(parent_dir, "old_scans")

		matches_found = False
		for filename in os.listdir(scans_dir):
			filepath = scans_dir + os.sep + filename
			self.logger.info(f"Found file: {filepath}")
			#
			scans_list = None
			if filepath[-4:] == ".csv":
				scans_list = self.load_csv(filepath)
			elif filepath[-5:] == ".xlsx":
				scans_list = self.load_xlsx(filepath)
			#
			if scans_list:
				scan_date = self.get_date_from_filename(filename)
				match_count = self.match_scans(scans_list, scan_date)
				if match_count > 0:
					matches_found = True
					new_filepath = old_scans_dir + os.sep + filename
					os.rename(filepath, new_filepath)
		if matches_found:	
			self.orders_storage.save()
			self.scans_storage.save()
			return matches_found
	
	def update_google_sheet(self):
		for customer in self.config.data["supported_customers"]:
			customer = self.config.data["supported_customers"][customer]
			spreadsheet_id = customer["google_spreadsheet_id"]
			data_sheet_data = [order for order in self.orders_storage.data if order.customer_id == customer["3PLC_customer_id"]]
			data_sheet_range = customer["google_sheet_data_range"]
			daily_summary_sheet_range = customer["google_sheet_daily_summary_range"]
			daily_summary_sheet_data = self.make_daily_orders_summary(data_sheet_data)
			weekly_summary_sheet_range = customer["google_sheet_weekly_summary_range"]
			weekly_summary_sheet_data = self.make_weekly_orders_summary(daily_summary_sheet_data)
			result_1 = self.google_api.update(spreadsheet_id=spreadsheet_id, range=data_sheet_range, values=self.orders_list_to_csv(data_sheet_data))
			result_2 = self.google_api.update(spreadsheet_id=spreadsheet_id, range=daily_summary_sheet_range, values=self.orders_summary_to_csv(daily_summary_sheet_data))
			result_3 = self.google_api.update(spreadsheet_id=spreadsheet_id, range=weekly_summary_sheet_range, values=self.orders_summary_to_csv(weekly_summary_sheet_data))
			if result_1.get("updates", {}).get("updatedRows", 0) == 0 or result_2.get("updates", {}).get("updatedRows", 0) == 0 or result_3.get("updates", {}).get("updatedRows", 0) == 0:
				self.logger.error(f"Error uploading summary to Google Sheets")
				self.logger.error(str(result_1))
				self.logger.error(str(result_2))
				self.logger.error(str(result_3))

	def orders_list_to_csv(self, orders) -> list:
		out = []
		#
		header_line = []
		for key in orders[0].__dict__.keys(): #keys of first element are converted to column names
			header_line.append(key)
		out.append(header_line)
		#
		for item in orders:
			out_line = []
			for value in item.__dict__.values():
				out_line.append(value)
			out.append(out_line)
		return out

	def make_daily_orders_summary(self, orders_list):
		start_date = datetime.strptime(self.config.data["program_start_date"][:10], "%Y-%m-%d")
		end_date = datetime.strptime(today(), "%Y-%m-%d")
		summary_dataset = {}
		for order in orders_list:
			date_index = start_date
			while date_index < end_date:
				date_str = datetime.strftime(date_index, "%Y-%m-%d")
				if order.creation_date[:10] == date_str:
					if not summary_dataset.get(date_str, False):
						summary_dataset.update(
							{
								date_str: {
									"date": order.creation_date[:10],
									"created_count": 0,
									"closed_count": 0,
									"printed_count": 0,
									"shipped_count": 0,
									"shipped_in_five_days": 0,
									"_days_to_ship_dataset": []
								}
							}
						)
					if order.creation_date:
						summary_dataset[date_str]["created_count"] += 1
					if order.close_date:
						summary_dataset[date_str]["closed_count"] += 1
					if order.print_date:
						summary_dataset[date_str]["printed_count"] += 1
					if order.ship_date:
						summary_dataset[date_str]["shipped_count"] += 1
						def days_to_ship(order):
							order_creation_date = datetime.strptime(order.creation_date[:10], "%Y-%m-%d")
							if order_creation_date.toordinal() % 7 == 6: #If date is Saturday
								order_creation_date = order_creation_date + timedelta(days=2)
							elif order_creation_date.toordinal() % 7 == 0: #If date is Saturday
								order_creation_date = order_creation_date + timedelta(days=1)

							order_ship_date = datetime.strptime(order.ship_date[:10], "%Y-%m-%d")
							if order_ship_date.toordinal() % 7 == 6: #If date is Saturday
								order_ship_date = order_ship_date + timedelta(days=2)
							elif order_ship_date.toordinal() % 7 == 0: #If date is Saturday
								order_ship_date = order_ship_date + timedelta(days=1)

							index = order_creation_date
							days_to_ship = 0
							while index < order_ship_date:
								if not index.toordinal() % 7 == 0 or index.toordinal() % 7 == 6:
									days_to_ship += 1
								index += timedelta(days=1)
							
							#self.logger.info(f"{order.creation_date} {order_creation_date}\n{order.ship_date} {order_ship_date}\n{str(days_to_ship)}")
							return days_to_ship
							
						days_to_ship = days_to_ship(order)
						summary_dataset[date_str]["_days_to_ship_dataset"].append(days_to_ship)
						if days_to_ship <= 5:
							summary_dataset[date_str]["shipped_in_five_days"] += 1
				date_index += timedelta(days=1)

		summary = []
		for date_str in summary_dataset:
			day = summary_dataset[date_str]
			day["average_days_to_ship"] = math.floor(sum([x for x in day["_days_to_ship_dataset"]])) / max(1, len(day["_days_to_ship_dataset"]))
			day["percent_shipped"] = day["shipped_count"] / day["created_count"]
			day["percent_shipped_in_5"] = day["shipped_in_five_days"] / max(1, day["shipped_count"])
			del day["_days_to_ship_dataset"]
			summary.append(day)
		self.logger.info(json.dumps(summary, indent=2))
		return summary

	def make_weekly_orders_summary(self, daily_summary):
		program_start_date = datetime.strptime(self.config.data["program_start_date"][:10], "%Y-%m-%d")
		program_start_date_prior_monday = datetime.fromordinal(math.floor((program_start_date - timedelta(days=1)).toordinal() / 7) * 7) + timedelta(days=1)
		index = program_start_date_prior_monday
		weekly_summary_dataset = {}
		for day in daily_summary:
			monday = datetime.fromordinal(math.floor((datetime.strptime(day["date"], "%Y-%m-%d") - timedelta(days=1)).toordinal() / 7) * 7) + timedelta(days=1)
			monday_str = datetime.strftime(monday, "%Y-%m-%d")
			if not weekly_summary_dataset.get(monday_str, False):
				weekly_summary_dataset.update(
					{
						monday_str: {
							"date": datetime.strftime(monday, "%Y-%m-%d"),
							"created_count": 0,
							"closed_count": 0,
							"printed_count": 0,
							"shipped_count": 0,
							"shipped_in_five_days": 0
						}
					}
				)
			weekly_summary_dataset[monday_str]["created_count"] += day["created_count"]
			weekly_summary_dataset[monday_str]["closed_count"] += day["closed_count"]
			weekly_summary_dataset[monday_str]["printed_count"] += day["printed_count"]
			weekly_summary_dataset[monday_str]["shipped_count"] += day["shipped_count"]
			weekly_summary_dataset[monday_str]["shipped_in_five_days"] += day["shipped_in_five_days"]
		#
		for day in weekly_summary_dataset.values():
			day["percent_shipped"] = day["shipped_count"] / day["created_count"]
			day["percent_shipped_in_5"] = day["shipped_in_five_days"] / max(1, day["shipped_count"])
		#
		weekly_summary = [x for x in weekly_summary_dataset.values()]
		self.logger.info(json.dumps(weekly_summary, indent=2))
		return [x for x in weekly_summary_dataset.values()]

	def orders_summary_to_csv(self, summary):
		out = []
		#
		header_line = []
		for key in summary[0].keys(): #keys of first element are converted to column names
			header_line.append(key)
		out.append(header_line)
		#
		for item in summary:
			out_line = []
			for value in item.values():
				out_line.append(value)
			out.append(out_line)
		return out

	def load_csv(self, filepath: str) -> list:
		scans_list = []
		with open(filepath, "r") as file:
			csv_reader = csv.reader(file, delimiter=",")
			for row in csv_reader:
				for scan in row:
					scan = str(scan)
					if self.looks_like_scan(scan):
						scans_list.append(scan)
		return scans_list

	def load_xlsx(self, filepath: str) -> list:
		scans_list = []
		worksheet = pyxl.load_workbook(filepath).active
		for row in worksheet.values:
			for scan in row:
				scan = str(scan)
				if self.looks_like_scan(scan):
					scans_list.append(scan)
		return scans_list

	def match_scans(self, scans_list: list, scan_date: str) -> int:
		match_count = 0
		expected_matches = len(scans_list)
		#
		for order in self.orders_storage.data:
			for scan in scans_list:
				if str(order.order_id) == scan:
					match_count += 1
					order = self.set_to_shipped(order, scan_date)
					if scan not in self.scans_storage.data:
						self.scans_storage.data.append(scan)
		self.logger.info(f"Found {str(match_count)} matches, expected {str(expected_matches)}")
		return match_count

	def set_to_shipped(self, order, scan_date=None):
		order.ship_status = "shipped"
		if scan_date == None:
			scan_date = today()
		#
		if order.ship_date != None:
			order.ship_date = min(order.ship_date, scan_date)
		else:
			order.ship_date = scan_date
		return order

	def looks_like_scan(self, scan) -> bool:
		scan = str(scan)
		if len(scan) == 6 or len(scan) == 8:
			if all([x in "1234567890" for x in scan]):
				return True
		return False

	def get_date_from_filename(self, filename: str) -> str:
		'''Quick and dirty. Defaults to today'''
		def is_year(part):
			if all([x in "1234567890" for x in part]) and len(part) == 4:
				if int(part) > 2000:
					return part
			return None
		
		def is_month(part):
			months = {
				"01": "jan",
				"02": "feb",
				"03": "mar",
				"04": "apr",
				"05": "may",
				"06": "jun",
				"07": "jul",
				"08": "aug",
				"09": "sep",
				"10": "oct",
				"11": "nov",
				"12": "dec"
			}
			for month_2_digits, month_name in months.items():
				if part.lower()[:3] == month_name:
					return month_2_digits
			return None

		def is_day(part):
			if all([x in "1234567890" for x in part]) and len(part) <= 2:
				if int(part) >= 1 and int(part) <= 31:
					return part.ljust(2, "0")
			return None
				
		year, month, day = None, None, None
		for part in filename.replace(".", "  ").split():
			if is_day(part):
				day = is_day(part)
			if is_month(part):
				month = is_month(part)
			if is_year(part):
				year = is_year(part)

		if all([year, month, day]):
			today_str = f"{year}-{month}-{day}"
			self.logger.info(f"Parsed date {today_str} from {filename}")
			return today_str
		else:
			today_str = today()
			self.logger.info(f"Unable to parse date from {filename}. Defaulting to {today_str}")
			return today_str

class Storage(Loggable):
	'''Allows access to data stored on disk'''
	def __init__(self, filepath: str, default_value=None):
		'''
		param filepath: path to .json file
		param default_value: default value for self.data
		'''
		if filepath[-5:] != ".json":
			raise AttributeError("Filepath must point to a valid .json file")
		self.filepath = filepath
		self.default_value = default_value
		self.data = self.load()

	def load(self) -> dict:
		try:
			with open(self.filepath, "r") as file:
				self.data = json.load(file, object_hook=MyJSONEncoder.object_hook)
		except FileNotFoundError:
			self.data = self.default_value
			with open(self.filepath, "x") as file:
				json.dump(self.data, file)
		return self.data

	def save(self):
		with open(self.filepath, "w") as file:
			json.dump(self.data, file, indent=4, cls=MyJSONEncoder)

	def data(self) -> object:
		'''Returns the object stored by Storage'''
		return self.data

class StoredList(Storage, Loggable):
	'''Extends Storage to function as a list with item indexes (for sorting)'''
	#Note: generalize to remove dependency on Storage?
	def __init__(self, filepath, default_value=None, default_index_function=None):
		'''
		param filepath: see Storage
		param default_value: see Storage
		param default_index_function: called when trying to add() to StoredList without an index param. Defaults to next_index()
		'''
		if default_value == None:
			default_value = []
		if default_index_function == None:
			self.index_counter = -1
			self.index_function = self.next_index
		super().__init__(filepath, default_value)
	
	def add(self, value, index=None):
		'''
		param value: value to be added
		param index: index to be added to value. Defaults to self.index_function() if missing
		'''
		if index == None:
			index = self.index_function()
		value._index = index
		if value not in self.data:
			self.data.append(value)

	def remove(self, value):
		if value in self.data:
			del value

	def next_index(self) -> int:
		self.index_counter += 1
		return self.index_counter

class Order:
	def __init__(self, data: dict):
		for key, value in data.items():
			self.__dict__[key] = value
	
	def __str__(self):
		return str(self.__dict__)

class MyJSONEncoder(json.JSONEncoder):
	def default(self, obj):
		if isinstance(obj, Order):
			obj.__dict__["__order__"] = True
			return obj.__dict__
		return json.JSONEncoder.default(self, obj)

	def object_hook(values):
		if "__order__" in values:
			return Order(values)
		return values

class WMS_API(Loggable):
	'''Gets orders data from 3PLC'''
	BYTES_USED = 0

	def __init__(self, config: Storage):
		'''
		params config: json file with config information
		'''
		self.config = config
		self.token = config.data["token"]
		token = self.get_token()
		if token:
			self.logger.info("3PLC access token refreshed")

	def get_token(self) -> dict:
		'''Returns a valid 3PLC access token, refreshing if needed'''
		creation_time = datetime.strptime(self.token.get("creation_time"), "%Y-%m-%d %H:%M:%S")
		token_duration = timedelta(seconds = self.token["contents"]["expires_in"])
		if datetime.now() > (creation_time + token_duration):
			token = self._refresh_token()
			if token:
				self.config.data["token"]["contents"] = token
				self.config.data["token"]["creation_time"] = datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M:%S")
				self.config.save()
				self.token = self.config.data["token"]
		return self.token
	
	def _refresh_token(self) -> dict:
		host_url = "https://secure-wms.com/AuthServer/api/Token"
		headers = {
			"Content-Type": "application/json; charset=utf-8",
			"Accept": "application/json",
			"Host": "secure-wms.com",
			"Accept-Language": "Content-Length",
			"Accept-Encoding": "gzip,deflate,sdch",
			"Authorization": "Basic " + self.config.data["auth_key"]
		}
		payload = json.dumps({
			"grant_type": "client_credentials",
			"tpl": self.config.data["tpl"],
			"user_login_id": self.config.data["user_login_id"]
		})

		response = requests.request("POST", host_url, data = payload, headers = headers, timeout = 3.0)
		self.log_data_usage(response) #TODO decorator
		if response.status_code == 200: #HTTP 200 == OK
			return response.json()
		else:
			self.logger.error("Unable to refresh token")
			self.logger.error(response.text)
		return None

	def get_order(self, order_id):
		order_id = str(order_id)
		raw_order = self._fetch_order(order_id)
		if raw_order:
			order = self._parse_order(raw_order)
			return order
		return None

	def _fetch_order(self, order_id: str):
		host_url = f"https://secure-wms.com/order/{order_id}"
		headers = {
			"Content-Type": "application/json; charset=utf-8",
			"Accept": "application/json",
			"Host": "secure-wms.com",
			"Accept-Language": "Content-Length",
			"Accept-Encoding": "gzip,deflate,sdch",
			"Authorization": "Bearer " + self.config.data["token"]["contents"]["access_token"]
		}
		response = requests.request("GET", host_url, data = {}, headers = headers, timeout = 30.0)
		self.log_data_usage(response) #TODO decorator
		return response

	def _parse_order(self, order: dict) -> Order:
		_order = None
		try:
			_order = Order({
				"order_id": order["ReadOnly"]["OrderId"],
				"batch_id": order["ReadOnly"].get("BatchIdentifier", {}).get("Id", None),
				"reference_id": order["ReferenceNum"],
				"creation_date": order["ReadOnly"]["CreationDate"],
				"close_date": order["ReadOnly"].get("ProcessDate", None),
				"print_date": order["ReadOnly"].get("PickTicketPrintDate", None),
				"customer_name": order["ReadOnly"]["CustomerIdentifier"]["Name"],
				"customer_id": order["ReadOnly"]["CustomerIdentifier"]["Id"],
				"carrier": order["RoutingInfo"].get("Carrier", None),
				"tracking_number": order.get("RoutingInfo", {}).get("TrackingNumber", ""),
				"consignee_name": order.get("ShipTo", {}).get("Name", None),
				"consignee_address_line": order["ShipTo"]["Address1"] + " " + order["ShipTo"].get("Address2", ""),
				"consignee_country": order["ShipTo"]["Country"],
				"consignee_city": order["ShipTo"]["City"],
				"consignee_state_province": order["ShipTo"].get("State", None),
				"consignee_postal_code": order["ShipTo"]["Zip"],
				"ship_status": None,
				"ship_date": None
			})
		except:
			self.logger.warning(f"Unable to parse order {order}")
			self.logger.warning(traceback.format_exc())
		return _order

	def get_3PLC_orders_since_date(self, customer_id: str, start_date: str, end_date: str=None) -> list:
		if end_date == None:
			end_date = now()
		if self.get_token():
			raw_orders = self._fetch_3PLC_orders_since_date(customer_id, start_date, end_date)
			if raw_orders:
				clean_orders = []
				for raw_order in raw_orders:
					result = self._parse_order(raw_order)
					if result:
						clean_orders.append(result)
				self.logger.info(f"Retreived {str(len(clean_orders))} new orders from 3PLC")
				return clean_orders
		return None

	def _fetch_3PLC_orders_since_date(self, customer_id: str, start_date: str, end_date: str) -> list:
		rql = f"readonly.CreationDate=gt={start_date};readonly.CreationDate=lt={end_date};readonly.customeridentifier.id=={customer_id}"
		max_pages = 1
		orders_list = []
		def _get_orders(pgnum):
			host_url = f"https://secure-wms.com/orders?pgsiz=1000&pgnum={pgnum}&rql={rql}&detail=Contacts"
			headers = {
				"Content-Type": "application/json; charset=utf-8",
				"Accept": "application/json",
				"Host": "secure-wms.com",
				"Accept-Language": "Content-Length",
				"Accept-Encoding": "gzip,deflate,sdch",
				"Authorization": "Bearer " + self.config.data["token"]["contents"]["access_token"]
			}
			response = requests.request("GET", host_url, data = {}, headers = headers, timeout = 30.0)
			self.log_data_usage(response) #TODO decorator
			return response
		response = _get_orders(pgnum=0)
		if response.status_code == 200:
			total_results = response.json()["TotalResults"]
			self.logger.info(f"{total_results} results found for {customer_id} from {start_date} to {end_date}")
			#
			page_orders = response.json()["ResourceList"]
			orders_list += page_orders
			#
			max_pages = math.ceil(total_results / 1000)
			if max_pages > 1:
				for page_count in range(2, max_pages+1):
					self.logger.info(f"Fetching page {str(page_count)} found for {customer_id} from {start_date} to {end_date}")
					response = _get_orders(pgnum=page_count)
					page_orders = response.json()["ResourceList"]
					orders_list += page_orders
			return orders_list
		else:
			self.logger.warning(f"Unable to fetch orders from {start_date} to {end_date} for {customer_id}")
			self.logger.warning(response.text)
			self.logger.warning(traceback.format_exc())
			return None

	def log_data_usage(self, request):
		'''Cuz higher-ups see the API usage bill
		Current estimate is ~2kb for a request (w/only Contacts)
		'''
		method_len = len(request.request.method)
		url_len = len(request.request.url)
		headers_len = len('\r\n'.join('{}{}'.format(key, value) for key, value in request.request.headers.items()))
		body_len = len(request.request.body if request.request.body else [])
		text_len = len(request.text)
		approx_len = method_len + url_len + headers_len + body_len + text_len
		self.logger.info(f"Used approx. {str(approx_len)} bytes")
		self.config.data["approx_bytes_used"] += approx_len
		return approx_len

class GoogleSheets_API(Loggable):
	SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

	def __init__(self, config: Storage):
		self.config = config
		creds = None
		if os.path.exists('./resources/token.pickle'):
			with open('./resources/token.pickle', 'rb') as token:
				creds = pickle.load(token)
		if not creds or not creds.valid:
			if creds:
				try: #TODO rework FOC to not do this
					creds.refresh(Request())
				except:
					os.remove('./resources/token.pickle')
					flow = InstalledAppFlow.from_client_secrets_file('./resources/credentials.json', self.SCOPES)
					creds = flow.run_local_server(port=0)
					creds.refresh(Request())
			else:
				flow = InstalledAppFlow.from_client_secrets_file('./resources/credentials.json', self.SCOPES)
				creds = flow.run_local_server(port=0)
			with open('./resources/token.pickle', 'wb') as token:
				pickle.dump(creds, token)

		service = build('sheets', 'v4', credentials=creds)
		self.sheet = service.spreadsheets()
	
	def update(self, spreadsheet_id: str, range: str, values):
		self.sheet.values().clear(spreadsheetId=spreadsheet_id, range=range).execute()
		body = {
			"values": values
		}
		result = self.sheet.values().append(spreadsheetId=spreadsheet_id, range=range, valueInputOption = "RAW", body = body).execute()
		if result.get("updates", {}).get("updatedRows", 0) == 0:
			self.logger.warning(f"Error updating Google Sheets:")
			self.logger.warning(result)
		return result
		
def init_logging():
	logger = logging.getLogger()
	logger.setLevel(logging.INFO)
	#
	file_handler = logging.FileHandler("./resources/log.txt")
	file_handler.setLevel(logging.ERROR)
	#
	console_handler = logging.StreamHandler()
	console_handler.setLevel(logging.DEBUG)
	#
	formatter = logging.Formatter(fmt="%(asctime)s %(levelname)s %(message)s", datefmt="%Y%m%d%H%M%S")
	file_handler.setFormatter(formatter)
	console_handler.setFormatter(formatter)
	#
	logger.addHandler(file_handler)
	logger.addHandler(console_handler)
	#
	logger.info("init CharmedTracker_V3")

def today():
	return datetime.strftime(datetime.now(), "%Y-%m-%d")

def now():
	return datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M:%S")

if __name__ == "__main__":
	init_logging()
	CharmedTracker().main()
	

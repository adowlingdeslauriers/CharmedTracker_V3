#CharmedTracker_V3.py
# Default Packages
from distutils import extension
import json
from datetime import datetime
import pathlib
import unittest
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

	def main(self):
		self.update_orders_list()
		matches_found = self.process_scans_folder()
		if matches_found or True:
			self.update_google_sheet()
	
	def update_orders_list(self):
		start_date = self.config.data["last_run_date"]
		end_date = now()
		for customer in self.config.data["supported_customers"]:
			customer_id = str(self.config.data["supported_customers"][customer])
			orders_list = self.wms_api.get_3PLC_orders_since_date(customer_id, start_date, end_date)
			if orders_list:
				orders_count = str(len(orders_list))
				self.logger.info(f"{orders_count} results parsed")
				#
				for order in orders_list:
					if "cancel" in order.reference_id.lower():
						order = self.set_to_shipped(order)
						self.logger.info(f"Order {order.order_id} reference_id: {order.reference_id} set to shipped")
					if "letter" in order.tracking_number.lower():
						order = self.set_to_shipped(order)
						self.logger.info(f"Order {order.order_id} {order.tracking_number} set to shipped")
					self.orders_storage.add(order, index=order.close_date)
				#
				self.orders_storage.save()
		self.config.data["last_run_date"] = now()
		self.config.save()
	
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
		sheet_id = self.config.data["google_sheet_id"]
		sheet_range = self.config.data["google_sheet_range"]
		self.google_api.update(sheet_id=sheet_id, range=sheet_range, values=self.orders_storage.data)

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
	'''Allows access to data stored on disk

	Example Usage:
	my_storage = Storage("C:/file.json")
	my_data = my_storage.data
	my_data.append(my_object)
	my_storage.save()
	'''
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

def today():
	return datetime.strftime(datetime.now(), "%Y-%m-%d")

def now():
	return datetime.strftime(datetime.now(), "%Y-%m-%d %H:%M:%S")

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
	BYTES_USED = 0
	'''Gets orders data from 3PLC'''
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

	def _parse_order(self, order: dict) -> Order:
		_order = None
		try:
			_order = Order({
				"order_id": order["ReadOnly"]["OrderId"],
				"batch_id": order["ReadOnly"].get("BatchIdentifier", {}).get("Id", None),
				"reference_id": order["ReferenceNum"],
				"creation_date": order["ReadOnly"]["CreationDate"],
				"close_date": order["ReadOnly"].get("ProcessDate", None),
				"print_date": order["ReadOnly"].get("pickTicketPrintDate", None),
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
				return clean_orders
		return None

	def _fetch_3PLC_orders_since_date(self, customer_id: str, start_date: str, end_date: str) -> list:
		rql = f"readonly.processDate=gt={start_date};readonly.processDate=lt={end_date};readonly.customeridentifier.id=={customer_id};readonly.isclosed==True"
		max_pages = 1
		orders_list = []
		def _get_orders(pgnum):
			host_url = f"https://secure-wms.com/orders?pgsiz=1000&pgnum={pgnum}&rql={rql}&detail=Contacts"
			#host_url = f"https://secure-wms.com/orders?pgsiz=1000&pgnum={pgnum}&rql={rql}"
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
			for page_count in range(1, max_pages):
				response = _get_orders(self, pgnum=page_count)
				page_orders = response.json()["ResourceList"]
				orders_list += page_orders
			return orders_list
		else:
			self.logger.warning(f"Unable to fetch orders from {start_date} to {end_date} for {customer_id}")
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
		self.logger.info(request.text)
		self.logger.info(f"Used approx. {str(approx_len)} bytes")
		self.config.data["approx_bytes_used"] += approx_len
		return approx_len

class GoogleSheets_API(Loggable):
	'''
	{'spreadsheetId': '1Cp9xkldWyeK5fyWK0QovsH57E1vgX13YqoSN638HsGI', 'updates': {'spreadsheetId': '1Cp9xkldWyeK5fyWK0QovsH57E1vgX13YqoSN638HsGI', 'updatedRange': 'MAIN!A1:L55027', 'updatedRows': 55027, 'updatedColumns': 12, 'updatedCells': 654052}}
	'''
	SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

	def __init__(self, config: Storage):
		self.config = config
		creds = None
		'''
		if os.path.exists('./resources/token.pickle'):
			with open('./resources/token.pickle', 'rb') as token:
				creds = pickle.load(token)
		if not creds or not creds.valid:
			if creds:
				creds.refresh(Request())
			else:
				flow = InstalledAppFlow.from_client_secrets_file('./resources/credentials.json', self.SCOPES)
				creds = flow.run_local_server(port=0)
			with open('./resources/token.pickle', 'wb') as token:
				pickle.dump(creds, token)
		'''
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
	
	def update(self, sheet_id: str, range: str, values):
		self.sheet.values().clear(spreadsheetId=sheet_id, range=range).execute()
		body = {
			"values": self.dict_to_csv(values)
		}
		result = self.sheet.values().append(spreadsheetId=sheet_id, range=range, valueInputOption = "RAW", body = body).execute()
		self.logger.info(result)
		return result

	def dict_to_csv(self, dict) -> list:
		out = []
		#
		header_line = []
		for key in dict[0].__dict__.keys(): #keys of first element are converted to column names
			header_line.append(key)
		out.append(header_line)
		#
		for item in dict:
			out_line = []
			for value in item.__dict__.values():
				out_line.append(value)
			out.append(out_line)
		return out
		
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

if __name__ == "__main__":
	init_logging()
	ct = CharmedTracker().main()
	

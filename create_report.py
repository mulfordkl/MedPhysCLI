import configparser
import os
import click
import sqlite3
import datetime
import openpyxl

if os.path.exists(os.path.join(os.path.dirname(__file__),'private')):
	data_dir = 'private'
else:
	data_dir = 'sample'

config = configparser.ConfigParser()
config.read(os.path.join(os.path.dirname(__file__), data_dir, 'configuration.ini'))

header_mapping = {
	"site": "B3",
	"location_detail": "B4",
	"unit_type": "B5",
	"mfr_model": "B6",
	"id_number": "B7",
	"report_type": "B9",
	"survey_date": "G3",
	"tested_by": "G4",
	"tested_by_spn": "G5",
	"checked_by": "G6",
	"checked_by_spn": "G7",
	"detector_model": "G8",
	"detector_sn": "G9",
	"detector_cal_date": "G10"
}

@click.command()
@click.argument("IDNumber")
@click.option('--type', default='Annual', help='The type of survey performed. Annual (Default), Other, Acceptance, or ACR')
@click.option('--date', default='today', help='pass the date of the survey as MM-DD-YYYY')
@click.option('--mod', default='', help='Any modifying text to be appended to the report filename')
def cli(idnumber, type, date, mod):

	report_date = get_report_date(date)
	if not report_date:
		return

	if type == "Acceptance":
		new_unit_info = new_unit_prompt()
		

		report_unit = Unit(new_unit_info)
		report_generator = ReportGenerator(report_date, type, mod)

		status = report_builder(report_unit, report_generator)
		if status is not None:
			print(status)
		else:
			return
	else:
		connection = connect_to_db(os.path.join(os.path.dirname(__file__), data_dir, 'db', 'equipment.db'))
		if connection is not None:
			results = select_equipment_by_id(connection, idnumber)
			results_length = len(results)
			if results_length == 0:
				print("ERROR - That ID number does not match any of our records")
				return
			for unit in results:

				report_unit = Unit(unit)
				report_generator = ReportGenerator(report_date, type, mod)

				status = report_builder(report_unit, report_generator)
				if status is not None:
					print(status)
				else:
					return
		else:
			print("Unable to connect to the database file - is it there?")
			return 

class Unit:
	def __init__(self, query_results):
		self.unit_id = query_results['id']
		self.unit_site = query_results['site']
		self.unit_location = query_results['location']
		self.unit_location_detail = query_results['location_detail']
		self.unit_type = self.type_list(query_results['type'])
		self.unit_mfr = query_results['manufacturer']
		self.unit_model = query_results['model']

	def type_list(self, query_type_results):
		unit_type = [query_type_results.replace(" ","-")]
		if unit_type[0] == "Rad/Fluoro":
			unit_type = ["X-Ray", "Fluoro"]
		return unit_type
	
	def get_template_type(self):
		mfr = self.unit_mfr
		template_type = []
		for unit_type in self.unit_type:
			unit_type = unit_type.lower()
			if unit_type in ["c-arm", "dental", "fluoro", "mini-c-arm", "o-arm", "x-ray"]:
				template_type.append(unit_type)
			elif unit_type == "portable-x-ray":
				if mfr in ["AGFA", "Samsung"]:
					template_type.append("portable_agfa")
				elif mfr == "GE":
					template_type.append("portable_amx")
			else:
				template_type.append(None)
		return template_type

	def print_type(self):
		print(self.unit_type)

class ReportGenerator:
	def __init__(self, date, report_type, mod):
		self.date = date
		self.report_type = report_type
		self.mod = mod

	def check_target_folder(self):
		base_folder = config['Dirs']['base_report_dir']
		month = "-".join([self.date.strftime("%m"), self.date.strftime("%B")])
		year = self.date.strftime("%Y")
		report_folder = os.path.join(base_folder, year, month)
		print("The report will go in folder " + report_folder)

		if os.path.exists(report_folder):
			print("This folder exists - no new directories created\n")
			
		else:
			print("This folder does not yet exist - making new directories")
			os.makedirs(report_folder, exist_ok=True)
		return report_folder
	
	def build_report_filenames(self, Unit, report_folder):
		report_filename_list = []
		for type in Unit.unit_type:
			report_filename = ""
			id = Unit.unit_id
			site = Unit.unit_site.replace(" ", "")
			unit_type = type.replace(" ","-")
			mfr = Unit.unit_mfr
			model = xstr(Unit.unit_model).replace(" ", "").replace("/","-")
			if model != "":
				mfr_model = '-'.join([mfr, model])
			else:
				mfr_model = mfr

			report_filename = "_".join([id, site, unit_type, mfr_model, self.report_type, self.date.strftime("%m-%d-%Y")])
			if self.mod:
				report_filename = "_".join([report_filename, self.mod])
			report_filename_list.append(os.path.join(report_folder, report_filename + ".xlsx"))

		return report_filename_list

	def get_template_files(self, template_types):
		template_filename_list = []
		for template_type in template_types:
			if template_type is None:
				template_filename_list.append(None)
			else:
				template_filename_list.append(os.path.join(os.path.dirname(__file__), data_dir, "templates", template_type + ".xlsx"))
		
		return template_filename_list

	def build_report(self, Unit, unit_type, template_file, report_filepath, header_mapping):
		try:
			template_wb = openpyxl.load_workbook(template_file)
		except:
			print("ERROR - some issue with opening the template")
			creation_status = 0
			return creation_status

		header_range = template_wb['Report']

		header_range[header_mapping["site"]] = Unit.unit_site
		header_range[header_mapping["location_detail"]] = xstr(Unit.unit_location) + " " + xstr(Unit.unit_location_detail)
		header_range[header_mapping["unit_type"]] = unit_type
		header_range[header_mapping["mfr_model"]] = Unit.unit_mfr + " " + xstr(Unit.unit_model)
		header_range[header_mapping["id_number"]] = Unit.unit_id

		header_range[header_mapping["report_type"]] = self.report_type
		header_range[header_mapping["survey_date"]] = self.date

		header_range[header_mapping["tested_by"]] = config["TestingInfo"]["tested_by"]
		header_range[header_mapping["tested_by_spn"]] = config["TestingInfo"]["tested_by_SPN"]
		header_range[header_mapping["checked_by"]] = config["TestingInfo"]["checked_by"]
		header_range[header_mapping["checked_by_spn"]] = config["TestingInfo"]["checked_by_SPN"]

		header_range[header_mapping["detector_model"]] = config["DetectorInfo"]["detector_model"]
		header_range[header_mapping["detector_sn"]] = config["DetectorInfo"]["detector_SN"]
		header_range[header_mapping["detector_cal_date"]] = config["DetectorInfo"]["detector_cal_date"]

		template_wb.save(report_filepath)
		creation_status = 1
		return creation_status

###############################################################################
## Database Access			 												 ##
###############################################################################

def connect_to_db(db_file):
	""" create a database connection to the SQLite database
		specified by db_file
	:param db_file: database file
	:return: Connection object or None
	"""
	conn = None
	if os.path.exists(db_file):
		try:
			conn = sqlite3.connect(db_file)
			return conn
		except:
			return conn
	else:
		return conn

def select_equipment_by_id(conn, id_number):
	"""
	Query tasks by priority
	:param conn: the Connection object
	:param priority:
	:return:
	"""
	cur = conn.cursor()
	cur.execute("SELECT * FROM equipment WHERE id=?", (id_number,))

	data = [dict((cur.description[i][0], value) for i, value in enumerate(row)) for row in cur.fetchall()]

	return data

###############################################################################
## Miscellaneous Functions   												 ##
###############################################################################

def new_unit_prompt():
	type_list = ["X-Ray", "Fluoro", "X-Ray/Fluoro", "Portable X-Ray", "C-Arm", "Mini C-Arm", "O-Arm"]
	
	new_unit_info = {}
	print("\nFill in the following form to add a new unit")
	new_unit_info["id"] = input("ID Number: ")
	new_unit_info['site'] = input("Site: ")
	new_unit_info['location'] = input("Location: ")
	new_unit_info['location_detail'] = input("Location detail (Nickname, Color/Number, Room): ")
	new_unit_info['type'] = input("Equipment Type:\n\tX-Ray\n\tFluoro\n\tX-Ray/Fluoro\n\tPortable X-Ray\n\tC-Arm\n\tMini C-Arm\n\tO-Arm\nEnter one of the above: ")
	while new_unit_info['type'] not in type_list:
		new_unit_info['type'] = input("Last input did not match valid types:\n\tX-Ray\n\tFluoro\n\tX-Ray/Fluoro\n\tPortable X-Ray\n\tC-Arm\n\tMini C-Arm\n\tO-Arm\nEnter one of the above: ")
	new_unit_info['manufacturer'] = input("Manufacturer: ")
	new_unit_info['model'] = input("Model: ")
	print('\nIs this info correct?')
	for prop in new_unit_info:
		print(prop,':',new_unit_info[prop])
	confirm_choice = input("\n[y/n]: ")
	print(new_unit_info)
	if confirm_choice in ["n", "no"]:
		print('\n')
		new_unit_info = new_unit_prompt()
		return new_unit_info
	else:
		print('\n')
		return new_unit_info

def get_report_date(date):
	if date == 'today':
			survey_date = datetime.date.today();
			return survey_date
	else: 
		date_split = date.split("-")
		report_month = int(date_split[0].lstrip("0"))
		report_day = int(date_split[1].lstrip("0"))
		report_year = int(date_split[2])
		try:
			survey_date = datetime.date(report_year, report_month, report_day)
			print(survey_date)
			return survey_date
		except:
			print("Date not entered in correct format")
			return

def report_builder(unit, report_generator):
	
	templates = unit.get_template_type()

	for template_type in templates:
		if template_type is None:
			status = "ERROR - That type of template does not exist in the template folder"
			return status


	report_dir_path = report_generator.check_target_folder();
	report_filenames = report_generator.build_report_filenames(unit, report_dir_path)

	for file in report_filenames:
		if not overwrite_check(file):
			status = "Cancelling report generation"
			return status

	template_filenames = report_generator.get_template_files(templates)
	num_types = len(unit.unit_type)
	for i in range(num_types):
		report_build_status = report_generator.build_report(unit, unit.unit_type[i], 
															template_filenames[i], 
															report_filenames[i], 
															header_mapping)
		if report_build_status == 0:
			print("Report not created - some error")
			status = "Error"
		elif report_build_status == 1:
			print("Report created at: {}".format(report_filenames[i]))
			status = None
	return status

def overwrite_check(report_filename):
	if os.path.exists(report_filename):
		check = input("That file exists - overwrite? [y/n]: ")
		if check in ['y', 'yes']:
			return True
		elif check in ['n', 'no']: 
			return False
		else:
			print("input not understood\n")
			overwrite_check(report_filename)
	else:
		return True


def xstr(s):
	if s is None:
		return ''
	return str(s)

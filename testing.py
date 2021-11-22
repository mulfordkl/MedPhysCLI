
class Report:

	def __init__(self, config, FNumber, Site, Location, MFR, Model, TestType):
		self.config = config
		self.FNumber = FNumber
		self.Site = Site
		self.Location = Location
		self.MFR = MFR
		self.Model = Model
		self.TestType = TestType

	def create_report(self):
		


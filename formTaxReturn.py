import pdf
from enum import auto, IntEnum
from PyPDF2 import PdfFileReader

class FormTaxReturn(pdf.Pdf):
	
	formtaxreturn_enum_extended = False
		
	def __init__(self, input_filepath = "<EMPTY>"):
		super().__init__(input_filepath)
		if(FormTaxReturn.formtaxreturn_enum_extended == False):
			extra_fields = [m.name for m in FormTaxReturn.Fields] + ['YEAR', 'EIN', 'NAME', 'TNAME', 'ADDRESS', 'FADDRESS']
			FormTaxReturn.Fields = IntEnum('FormTaxReturn.Fields', extra_fields, start = 0)
			FormTaxReturn.formtaxreturn_enum_extended = True

		for x in range((FormTaxReturn.Fields.FILENAME.value), (FormTaxReturn.Fields.FADDRESS.value)):
			self.list_fields.append("<EMPTY>")

	def displayEnum(self):
		for x in FormTaxReturn.Fields:
			print(x.value," ",x.name)

	def display(self):
			for x in range(0, len(self.list_fields)):
				print("FORMTXRET-FIELD-[", FormTaxReturn.Fields(x), "] -> ", self.list_fields[x])
			print("\n")

	def extractData(self):
		pdf_reader = PdfFileReader(open(self.list_fields[FormTaxReturn.Fields.FILENAME.value], "rb"))
		#add security checks here for opening pdf
		#what if this is called with empty filename

		dictionary = pdf_reader.getFormTextFields()
		list_of_dict_values = []

		for value in dictionary.values(): 
			list_of_dict_values.append(value)

		#extracts EIN
		for x in range(0 , 8):
			list_of_dict_values[0] = str(list_of_dict_values[0]) + str(list_of_dict_values[x+1])
		self.list_fields[FormTaxReturn.Fields.EIN.value] = list_of_dict_values[0]
		
		#extracts NAME
		self.list_fields[FormTaxReturn.Fields.NAME.value] = str(list_of_dict_values[9])
		
		#extracts TNAME
		self.list_fields[FormTaxReturn.Fields.TNAME.value] = str(list_of_dict_values[10])
		
		#extracts ADDRESS
		for x in range(11, 14):
			list_of_dict_values[11] = str(list_of_dict_values[11]) + "  " + str(list_of_dict_values[x+1])					
		self.list_fields[FormTaxReturn.Fields.ADDRESS.value] = list_of_dict_values[11]

		#extracts FADDRESS
		for x in range(15, 17):
			list_of_dict_values[15] = str(list_of_dict_values[15]) + "  " + str(list_of_dict_values[x+1])					
		self.list_fields[FormTaxReturn.Fields.FADDRESS.value] = list_of_dict_values[15]

		#replaces 'None' in list
		self.list_fields = [sub.replace("None", "") for sub in self.list_fields] 

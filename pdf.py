from enum import auto, IntEnum

class Pdf:

	class Fields(IntEnum):
		def _generate_next_value_(name, start, count, last_values):
			return count

		FILENAME = auto()
	
	def __init__(self, input_filename = "<EMPTY>"):
		self.list_fields = ["<EMPTY>"]
		self.list_fields[Pdf.Fields.FILENAME.value] = input_filename

	def display(self):
		for x in range(0, len(self.list_fields)):
			print("PDF--FIELD-[", Pdf.Fields(x), "] -> ", self.list_fields[x])
			print("\n")
	
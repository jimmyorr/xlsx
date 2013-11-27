#!/usr/bin/env python

import Tkinter
import codecs
import openpyxl.cell
import openpyxl.reader.excel
import optparse
import sys
import tkFileDialog
import tkMessageBox


def encode_str(s):
  return codecs.BOM_UTF16_BE + unicode(s).encode("utf_16_be")


def handle_strings(fdf_strings):
  for (key, value) in fdf_strings:
    yield "<<\n/V (%s)\n/T (%s)>>\n" % (encode_str(value), encode_str(key))


def handle_checkboxes(fdf_checkboxes):
  for (key, value) in fdf_checkboxes:
    if value:
      yield "<<\n/V/1\n/T (%s)>>\n" % (encode_str(key))
    else:
      yield "<<\n/V/0\n/T (%s)>>\n" % (encode_str(key))


def handle_booleans(fdf_booleans):
  for (key, value) in fdf_booleans:
    if value:
      yield "<<\n/V/X\n/T (%s)>>\n" % (encode_str(key))
    else:
      yield "<<\n/V/Off\n/T (%s)>>\n" % (encode_str(key))


def generate_fdf(fdf_strings=[], fdf_checkboxes=[], fdf_booleans=[]):
  fdf = ["%FDF-1.2\n%\xe2\xe3\xcf\xd3\r\n"]
  fdf.append("1 0 obj\n<<\n/FDF\n")

  fdf.append("<<\n/Fields [\n")
  fdf.append("".join(handle_strings(fdf_strings)))
  fdf.append("".join(handle_checkboxes(fdf_checkboxes)))
  fdf.append("".join(handle_booleans(fdf_booleans)))
  fdf.append("]\n")

  fdf.append(">>\n")
  fdf.append(">>\nendobj\n")
  fdf.append("trailer\n\n<<\n/Root 1 0 R\n>>\n")
  fdf.append("%%EOF\n\x0a")

  return "".join(fdf)


def column_index_from_string(index):
  if index.isdigit():
    return index
  else:
    return openpyxl.cell.column_index_from_string(index) - 1


TRUE_VALUES = (1, "1", "true", "True", "X", "x", True)


class Xlsx2fdf(object):
  def __init__(self):
    pass

  def validate(self):
    if (not self.input_xlsx or not self.sheet_name or not self.key_column or
        not self.value_column or not self.type_column or not self.output_fdf):
      return False
    else:
      return True

  def process(self):
    workbook = openpyxl.reader.excel.load_workbook(filename=self.input_xlsx)
    sheet = workbook.get_sheet_by_name(name=self.sheet_name)
    key_column = sheet.columns[column_index_from_string(self.key_column)]
    value_column = sheet.columns[column_index_from_string(self.value_column)]
    type_column = sheet.columns[column_index_from_string(self.type_column)]

    fdf_strings = []
    fdf_checkboxes = []
    fdf_booleans = []
    for index, key_cell in enumerate(key_column):
      value_cell = value_column[index]
      if key_cell.value and value_cell.value is not None:
        if str.lower(str(type_column[index].value)) == "checkbox":
          key_value_tuple = (key_cell.value,
                             bool(value_cell.value in TRUE_VALUES))
          fdf_checkboxes.append(key_value_tuple)
        elif str.lower(str(type_column[index].value)) == "boolean":
          key_value_tuple = (key_cell.value,
                             bool(value_cell.value in TRUE_VALUES))
          fdf_booleans.append(key_value_tuple)
        else:
          key_value_tuple = (key_cell.value, value_cell.value)
          fdf_strings.append(key_value_tuple)
    fdf = generate_fdf(fdf_strings, fdf_checkboxes, fdf_booleans)
    fdf_file = open(self.output_fdf, "wb")
    fdf_file.write(fdf)
    fdf_file.close()

  def set_input_xlsx_tk(self, input_xlsx_var):
    self.input_xlsx = tkFileDialog.askopenfilename(filetypes=[("xlsx files", "*.xlsx")])
    input_xlsx_var.set(self.input_xlsx)

  def set_output_fdf_tk(self, output_fdf_var):
    self.output_fdf = tkFileDialog.asksaveasfilename(filetypes=[("fdf file", "*.fdf")])
    output_fdf_var.set(self.output_fdf)


class Xlsx2fdfGui(object):
  def __init__(self, xlsx2fdf):
    self.xlsx2fdf = xlsx2fdf

    self.root = Tkinter.Tk()

    self.input_xlsx_var = Tkinter.StringVar()
    self.input_xlsx_var.set(self.xlsx2fdf.input_xlsx)

    self.sheet_name_var = Tkinter.StringVar()
    self.sheet_name_var.set(self.xlsx2fdf.sheet_name)

    self.key_column_var = Tkinter.StringVar()
    self.key_column_var.set(self.xlsx2fdf.key_column)

    self.value_column_var = Tkinter.StringVar()
    self.value_column_var.set(self.xlsx2fdf.value_column)

    self.type_column_var = Tkinter.StringVar()
    self.type_column_var.set(self.xlsx2fdf.type_column)

    self.output_fdf_var = Tkinter.StringVar()
    self.output_fdf_var.set(self.xlsx2fdf.output_fdf)

  def run(self):
    self.root.title('xlsx2fdf')

    frame = Tkinter.Frame(self.root, width=1000)
    frame.pack(fill=Tkinter.BOTH, expand=True)

    top_frame = Tkinter.Frame(self.root)
    top_frame.pack(fill=Tkinter.BOTH, expand=True)

    left_frame = Tkinter.Frame(top_frame, width=100)
    left_frame.pack(fill=Tkinter.BOTH, expand=False, side=Tkinter.LEFT)
    input_xlsx_button = Tkinter.Button(left_frame, text="Set --input_xlsx...",
        command=lambda: self.xlsx2fdf.set_input_xlsx_tk(self.input_xlsx_var))
    input_xlsx_button.pack(fill=Tkinter.BOTH, expand=True)
    sheet_name_button = Tkinter.Label(left_frame, text="Set --sheet_name")
    sheet_name_button.pack(fill=Tkinter.BOTH, expand=True)
    key_column_button = Tkinter.Label(left_frame, text="Set --key_column")
    key_column_button.pack(fill=Tkinter.BOTH, expand=True)
    value_column_button = Tkinter.Label(left_frame, text="Set --value_column")
    value_column_button.pack(fill=Tkinter.BOTH, expand=True)
    type_column_button = Tkinter.Label(left_frame, text="Set --type_column")
    type_column_button.pack(fill=Tkinter.BOTH, expand=True)
    output_fdf_button = Tkinter.Button(left_frame, text="Set --output_fdf...",
        command=lambda: self.xlsx2fdf.set_output_fdf_tk(self.output_fdf_var))
    output_fdf_button.pack(fill=Tkinter.BOTH, expand=True)

    right_frame = Tkinter.Frame(top_frame)
    right_frame.pack(fill=Tkinter.BOTH, expand=True)
    input_xlsx_entry = Tkinter.Entry(right_frame, textvariable=self.input_xlsx_var, state="readonly")
    input_xlsx_entry.pack(fill=Tkinter.BOTH, expand=True)
    sheet_name_entry = Tkinter.Entry(right_frame, textvariable=self.sheet_name_var)
    sheet_name_entry.pack(fill=Tkinter.BOTH, expand=True)
    key_column_entry = Tkinter.Entry(right_frame, textvariable=self.key_column_var)
    key_column_entry.pack(fill=Tkinter.BOTH, expand=True)
    value_column_entry = Tkinter.Entry(right_frame, textvariable=self.value_column_var)
    value_column_entry.pack(fill=Tkinter.BOTH, expand=True)
    type_column_entry = Tkinter.Entry(right_frame, textvariable=self.type_column_var)
    type_column_entry.pack(fill=Tkinter.BOTH, expand=True)
    output_fdf_entry = Tkinter.Entry(right_frame, textvariable=self.output_fdf_var, state="readonly")
    output_fdf_entry.pack(fill=Tkinter.BOTH, expand=True)

    bottom_frame = Tkinter.Frame(self.root)
    bottom_frame.pack(fill=Tkinter.BOTH, expand=True)

    process_button = Tkinter.Button(
        bottom_frame, text="Process", width=50,
        command=lambda: self.process_tk(self.sheet_name_var,
                                        self.key_column_var,
                                        self.value_column_var,
                                        self.type_column_var))
    process_button.pack()

    self.root.mainloop()

  def process_tk(self, sheet_name_var, key_column_var,
                 value_column_var, type_column_var):
    self.xlsx2fdf.sheet_name = sheet_name_var.get()
    self.xlsx2fdf.key_column = key_column_var.get()
    self.xlsx2fdf.value_column = value_column_var.get()
    self.xlsx2fdf.type_column = type_column_var.get()

    if not self.xlsx2fdf.validate():
      tkMessageBox.showwarning("Usage", "Missing or invalid value")
    else:
      try:
        self.xlsx2fdf.process()
      except Exception as e:
        tkMessageBox.showwarning("Error!", message=e)

def main(argv=None):
  usage = "usage: %prog [options]"
  parser = optparse.OptionParser(usage)
  parser.add_option("-i", "--input_xlsx", dest="input_xlsx",
      help="xlsx filename to read from")
  parser.add_option("-s", "--sheet_name", dest="sheet_name",
      help="name of sheet in xlsx file")
  parser.add_option("-k", "--key_column", dest="key_column",
      help="column in sheet containing pdf's keys")
  parser.add_option("-v", "--value_column", dest="value_column",
      help="column in sheet containing values")
  parser.add_option("-t", "--type_column", dest="type_column",
      help="column in sheet identifying the type of the value " +
      "(boolean, checkbox, or blank/string)")
  parser.add_option("-o", "--output_fdf", dest="output_fdf",
      help="fdf filename to write to")
  parser.add_option("-g", "--gui", dest="use_gui", action="store_true",
      help="use graphical user interface", default=True)
  parser.add_option("-n", "--nogui", dest="use_gui", action="store_false",
      help="do not use graphical user interface")

  if argv is None:
    argv = sys.argv
  options, unused_args = parser.parse_args()

  xlsx2fdf = Xlsx2fdf()
  xlsx2fdf.input_xlsx = options.input_xlsx
  xlsx2fdf.sheet_name = options.sheet_name
  xlsx2fdf.key_column = options.key_column
  xlsx2fdf.value_column = options.value_column
  xlsx2fdf.type_column = options.type_column
  xlsx2fdf.output_fdf = options.output_fdf

  if options.use_gui:
    xlsx2fdfGui = Xlsx2fdfGui(xlsx2fdf)
    xlsx2fdfGui.run()
  else:
    if not xlsx2fdf.validate():
      parser.print_help()
      return 1
    xlsx2fdf.process()


if __name__ == "__main__":
  sys.exit(main())

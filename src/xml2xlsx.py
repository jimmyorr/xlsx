#!/usr/bin/env python

from xml.dom import minidom, Node
import openpyxl
import sys
import tkFileDialog

def xml2xlsx(xml_filename):
  wb = openpyxl.Workbook()
  ws = wb.create_sheet(0, "mapping from xml")

  xmldoc = minidom.parse(xml_filename)

  for field in xmldoc.getElementsByTagName("field"):
    descriptors = [field.attributes["xfdf:original"].value]
    parent = field.parentNode
    while parent.nodeType == Node.ELEMENT_NODE:
      if parent.attributes.get("xfdf:original") is not None:
        descriptors.append(parent.attributes.get("xfdf:original").value)
      parent = parent.parentNode
    descriptors.reverse()
    ws.append([".".join(descriptors), field.childNodes[0].nodeValue])

  wb.save(xml_filename + ".xlsx")


def main(argv=None):
  if argv is None:
    argv = sys.argv

  if len(argv) > 1:
    xml_filename = argv[1]
  else:
    xml_filename = tkFileDialog.askopenfilename(filetypes=[("xml mapping files", "*.xml")])

  xml2xlsx(xml_filename)

if __name__ == "__main__":
  sys.exit(main())

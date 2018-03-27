#!/usr/bin/env python3
import sys
import xml.dom.minidom
import zipfile

def parse(raw_data): # return matrix
    strs = raw_data.split("\n")
    tbl = []
    for s in strs:
        if "<row" in s:
            tbl.append([])
            continue
        if not "<c" in s:
            continue
        pos = s.find("s=\"")
        value = int( s[pos + 3] )
        tbl[-1].append(value)
    return tbl


def to_html(tbl):
    tmpl = "{}"
    try:
        tmpl_file = open("template.html", "r")
        tmpl = tmpl_file.read()
        tmpl_file.close()
    except IOError:
        print("Can't open 'template.html'")
    htbl = "<table>\n"
    for tr in tbl:
        htbl += "   <tr>\n"
        for td in tr:
            htbl += "       <td class=\"s"+str(td)+"\"></td>\n"
        htbl += "   </tr>\n"
    htbl += "</table>"
    result = tmpl.replace("***table here***", htbl)
    return result

def extract(arch_name, target):
    result = ""
    try:
        arch = zipfile.ZipFile(arch_name, "r")
        result = arch.read(target)
        arch.close()
    except (IOError, KeyError):
        print("Incorrect archive provided")
    return result

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: {} <xslx or xml file>".format(sys.argv[0]))
        exit()
    filename = sys.argv[1]
    base = filename.split(".")[0]
    ext = filename.split(".")[-1]
    if len(filename.split(".")) < 2:
        print("Incorrect filename specified")
        exit()
    
    raw_data = ""
    if ext == "xlsx":
        raw_data = extract(filename, "xl/worksheets/sheet1.xml")
    elif ext == "xml":
        try:
            xml_file = open(filename, "r")
            raw_data = xml_file.read()
            xml_file.close()
        except IOError:
            print("Incorrect xml provided")
    else:
        print("Unsupported file provided")

    xml = xml.dom.minidom.parseString(raw_data)
    data = xml.toprettyxml()
    table = parse(data)
    html = to_html(table)

    try:
        output_file = open("{}.html".format(base), "w")
        output_file.write(html)
        output_file.close()
        print("Result written to '{}.html'".format(base))
    except IOError:
        print("Can't write result")
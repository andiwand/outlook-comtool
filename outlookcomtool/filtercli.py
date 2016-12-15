import gzip
import json
import csv
import re
import argparse

pattern = re.compile(r"^(.*)\:\s*?$", re.MULTILINE)

def parse_body(body):
    result = {}
    start = 0
    last = None
    while True:
        match = pattern.search(body, start)
        value_end = match.start() if match else None
        value = body[start:value_end].strip()
        if last: result[last] = value
        if not match: break
        start = match.end()
        last = match.group(1)
    return result

def extra_columns():
    return ["DU / SIE", "Frau / Herr", "Sprachen"]

def extra(contact):
    columns = extra_columns()
    result = [""] * len(columns)
    extra = parse_body(contact["Body"])
    for i, column in enumerate(columns):
        if column in extra: result[i] = extra[column]
    return result

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="outlook filter tool")
    parser.add_argument("-a", "--attributes", help="filter attributes (comma separated)", default="FullName,FirstName,LastName,Title,Suffix,Email1Address,Email2Address,Email3Address,HomeAddressCountry,BusinessAddressCountry")
    parser.add_argument("-i", "--input", help="input file", required=True)
    parser.add_argument("-o", "--output", help="output file", required=True)
    args = parser.parse_args()
    
    fin = gzip.open(args.input) if args.input.endswith(".gz") else open(args.input, "r")
    fout = open(args.output, "w")
    attributes = args.attributes.split(",")
    
    contacts = json.load(fin)
    fin.close()
    
    csvwriter = csv.writer(fout, delimiter=",", quotechar="\"", quoting=csv.QUOTE_ALL)
    
    csvwriter.writerow(attributes + extra_columns())
    for i, contact in enumerate(contacts):
        print(i, contact["FullName"])
        row = [contact[key] for key in attributes]
        row.extend(extra(contact))
        row = [cell.encode("utf-8") for cell in row]
        csvwriter.writerow(row)

    fout.close()
    
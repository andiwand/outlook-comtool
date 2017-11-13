#!/usr/bin/env python2

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

def extra(contact, extras):
    columns = extras
    result = [""] * len(columns)
    extra = parse_body(contact["Body"])
    for i, column in enumerate(columns):
        if column in extra: result[i] = extra[column]
    return result

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="outlook filter tool")
    parser.add_argument("-a", "--attributes", help="filter attributes (comma separated)", default="FullName,FirstName,LastName,Title,Suffix,HomeAddressCountry,BusinessAddressCountry,Categories,Email1Address,Email2Address,Email3Address")
    parser.add_argument("-e", "--extra", help="parse body for extra information (comma separated)", default="DU / SIE,Frau / Herr,Sprachen")
    parser.add_argument("-i", "--input", help="input file", required=True)
    parser.add_argument("-o", "--output", help="output file", required=True)
    args = parser.parse_args()

    fin = gzip.open(args.input) if args.input.endswith(".gz") else open(args.input, "r")
    fout = open(args.output, "w")
    attributes = args.attributes.split(",")
    extras = args.extra.split(",")

    contacts = json.load(fin)
    fin.close()

    csvwriter = csv.writer(fout, lineterminator="\n", delimiter=",", quotechar="\"", quoting=csv.QUOTE_ALL)

    csvwriter.writerow(attributes + extras)
    for i, contact in enumerate(contacts):
        print(i, contact["FullName"])
        row = [contact[key] for key in attributes]
        if extras: row.extend(extra(contact, extras))
        row = [cell.encode("utf-8") for cell in row]
        csvwriter.writerow(row)

    fout.close()

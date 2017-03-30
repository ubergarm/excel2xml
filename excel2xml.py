#!/usr/bin/env python3
import sys
import json
import argparse
import pandas as pd
import dicttoxml

# CLI arguments
parser = argparse.ArgumentParser(formatter_class=argparse.ArgumentDefaultsHelpFormatter,
                                 description='Convert Excel to XML. Defaults to stdin stdout.')
parser.add_argument('--infile',
                    type=str,
                    default='-',
                    help='filename, url, or "-" for stdin')
parser.add_argument('--outfile',
                    type=str,
                    default='-',
                    help='filename, or "-" for stdout')
parser.add_argument('--sheetnum',
                    type=int,
                    default=0,
                    help='number (0-indexed) of worksheet to read')
parser.add_argument('--sheetname',
                    type=str,
                    default=None,
                    help='exact text name of worksheet to read (overrides sheetnum)')
parser.add_argument('--skiprows',
                    type=int,
                    default=0,
                    help='number of rows to skip at the top of input')
parser.add_argument('--header',
                    type=int,
                    default=0,
                    help='row (0-indexed after skiprows) to use for the column labels')
parser.add_argument('--parsecols',
                    type=str,
                    default=None,
                    help='column range to parse e.g. "A,B,D:AF"')
parser.add_argument('--cdata',
                    dest='cdata',
                    action='store_true',
                    help='wrap all xml fields in CDATA tag')
parser.add_argument('--no-cdata',
                    dest='cdata',
                    action='store_false',
                    help='do NOT wrap all xml fields in CDATA tag')
parser.set_defaults(cdata=True)

args = parser.parse_args()

# map in/out files
infile = sys.stdin
outfile = sys.stdout
if args.infile != '-':
    infile = args.infile
if args.outfile != '-':
    try:
        outfile = open(args.outfile, 'w')
    except Exception as e:
        print(e)
        sys.exit(1)

# Convert EXCEL to DataFrame
df = pd.read_excel(infile,
                   sheetname=args.sheetname or args.sheetnum,
                   skiprows=args.skiprows,
                   header=args.header,
                   parse_cols=args.parsecols,
                   ignore_index=True).dropna(axis=1)

# Convert DataFrame to JSON
obj = json.loads(df.to_json(orient='records', force_ascii=True))

# Convert JSON to XML
xml = dicttoxml.dicttoxml(obj, attr_type=False, cdata=args.cdata)

# Save encoded results
outfile.write(xml.decode('utf-8'))

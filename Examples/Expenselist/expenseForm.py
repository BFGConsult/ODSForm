#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys, getopt, os, io, json
import ezodf, yaml
from datetime import date, datetime
from SpreadsheetMap import SpreadsheetMap

mapping = {   'UseFieldNamesDirectly': True,
    'mapping': {   'account': {   'cell': 'D10',
                                    'pattern': '\\d{4}\\.\\d{2}\\.\\d{5}|\\d{11}',
                                    'type': 'integer'},
                    'address': 'D8',
                    'entries': {   'cell': 'A15',
                                    'multiple': {   'entry': 'row',
                                                     'span': 11},
                                    'type': [   'date',
                                                 None,
                                                 'string',
                                                 None,
                                                 None,
                                                 None,
                                                 'integer']},
                    'name': 'd7',
                    'place': {   'type': 'string'},
                    'purpose': 'a4',
                    'zip': {   'pattern': '\\d{1,4}', 'type': 'integer'},
                    'zip-post': {   'cell': 'D9',
                                     'combine': ['zip', 'place'],
                                     'format': '%04d %s'}}}

#Uncomment below to get a sample map file
#import pprint
#pprint.PrettyPrinter(indent=4).pprint(mapping)
#print(json.dumps(mapping, sort_keys=True, indent=4))



def loadConfig(filename):
   try:
      if filename=='-':
         filename=sys.stdin
      conf = json.load(filename)
   except:
      stream = open(filename, "r")
      conf = yaml.load(stream)
      stream.close()
   return conf

class Utlegg:
   @staticmethod
   def getEntry(mydate, text, value):
      if mydate:
         mydate=datetime.strptime(mydate, '%Y-%m-%d').date()
      else:
         mydate=date.today()

      entry=[mydate, text, value]
      return entry

   def getEntryFromList(list):
      return getEntry(list[0], list[1], list[2])

   def fixEntry(obj):
      if type(obj) == type([]) :
         return getEntryFromList(obj)
      else:
         return obj

def usage():
   sys.stderr.write("Usage: %s [-o outputfile] inputfile <Purpose>\n\t\t< <Date1> <Text1> <Amount1> > [ <Date2> <Text2> <Amount2> ] […]\n" % sys.argv[0])

def custom_sort(t):
   t = t.upper()
   coord = SpreadsheetMap.getCoord(t)
   return coord[1]*100000+coord[0]

if __name__ == "__main__":
   try:
      opts, args = getopt.getopt(sys.argv[1:], "c:d:o:v",
                                 ["output=", "conffile=", "date=", "verbose"])
   except getopt.GetoptError:
      usage()
      sys.exit(2)

   outputfile = None

   datestamp = date.today()
   conffile = os.path.expanduser("staticdata.yml")
   verbose = False
   
   for o, a in opts:
      if o in ("-o", "--output"):
         outputfile = a
      if o in ("-c", "--conffile"):
         conffile = a
      if o in ("-d", "--date"):
         datestamp = datetime.strptime(a, '%Y-%m-%d').date()
      if o in ("-v", "--verbose"):
         verbose = True

   nargs=len(args)
   if ( (nargs - 2) % 3 != 0 ):
      usage()
      sys.exit(2)

   inputfile = args[0]
   if outputfile is None:
      outputfile = inputfile[inputfile.rfind('/') + 1:]
      outputfile = outputfile[:outputfile.rfind('.')] + ".ods"

   conf = loadConfig(conffile)

   if nargs > 1:
      conf['purpose'] = args[1]

   purpose = conf['purpose']
   
   if not purpose:
      usage()
      sys.exit(2)

   if 'entries' in conf:
      entries=conf['entries']
   else:
      entries=[]

   for i in range(len(entries)):
      entries[i]= fixEntry(entries[i])

   for i in range(0, (nargs-2)/3):
      dato=args[i*3+2]
      text=args[i*3+3].decode('utf8')
      verdi=float(args[i*3+4])
      entry=Utlegg.getEntry(dato, text, verdi)
      entries.append(entry)

   conf['entries'] = entries

   mapping['spreadsheet'] = inputfile
   smap = SpreadsheetMap(mapping,conf)

   smap.save(outputfile)
   if verbose:
      keys = sorted(smap.keys(), key=custom_sort)
      for key in keys:
         print "%s : %s" % (key, smap[key])

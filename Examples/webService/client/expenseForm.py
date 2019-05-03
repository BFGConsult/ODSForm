#!/usr/bin/python
# -*- coding: utf-8 -*-

from __future__ import print_function

import sys, getopt, os, json, yaml
from datetime import date, datetime

try:
    # For Python 3.0 and later
    from urllib.request import urlopen, Request
except ImportError:
    # Fall back to Python 2's urllib2
    from urllib2 import urlopen, Request

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
   print("Usage: %s [-o outputfile] inputfile <Purpose>\n\t\t< <Date1> <Text1> <Amount1> > [ <Date2> <Text2> <Amount2> ] […]\n" % sys.argv[0], file=sys.stderr)

if __name__ == "__main__":
   try:
      opts, args = getopt.getopt(sys.argv[1:], "c:d:o:u:v",
                                 ["output=", "conffile=", "date=",
                                  "url=", "verbose"])
   except getopt.GetoptError:
      usage()
      sys.exit(2)

   outputfile = None

   datestamp = date.today()
   conffile = os.path.expanduser("staticdata.yml")
   verbose = False
   url=None

   for o, a in opts:
      if o in ("-o", "--output"):
         outputfile = a
      if o in ("-c", "--conffile"):
         conffile = a
      if o in ("-u", "--url"):
         url = a
      if o in ("-d", "--date"):
         datestamp = datetime.strptime(a, '%Y-%m-%d').date()
      if o in ("-v", "--verbose"):
         verbose = True

   nargs=len(args)
   if ( (nargs - 1) % 3 != 0 ):
      usage()
      sys.exit(2)

   conf = loadConfig(conffile)

   curl = conf.pop('url', None)
   if not url:
      if not curl:
         print('No URL given', file=sys.stderr)
         sys.exit(1)
         
      url = curl
   
   if nargs > 1:
      conf['purpose'] = args[0]

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

   for i in range(0, (nargs-1)//3):
      dato=args[i*3+1]
      text=args[i*3+2]
      verdi=float(args[i*3+3])
      entry=Utlegg.getEntry(dato, text, verdi)
      entries.append(entry)

   conf['entries'] = entries

   req = Request(url)
   req.add_header('Content-Type', 'application/json')

   try:
       response = urlopen(req, json.dumps(conf, default=str).encode('utf-8'))
   except:
       print('Connection refused: %s' % (url), file=sys.stderr)
       EX_UNAVAILABLE=69
       sys.exit(EX_UNAVAILABLE)

   if (response.code==200):
       content = response.read()
       if not outputfile:
           cd = response.headers['content-disposition']
           if cd:
               outputfile = dict((elem.strip()+'=').split('=')[0:2]
                  for elem in cd.split(';')).get('filename','Output.ods')
       with open(outputfile, 'wb') as outfile:
           outfile.write(content)
   else:
       print('Unhandled response code %d' % (response.code), file=sys.stderr)
       sys.exit(1)

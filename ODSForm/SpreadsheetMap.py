#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

import ezodf

import copy, io, os, sys, re, pprint
from datetime import date, datetime

try:
  basestring = basestring
except NameError:
  basestring = str
try:
    FileNotFoundError = FileNotFoundError
except NameError:
    FileNotFoundError = IOError

#cellnumber = re.compile('([A-Z]{1,2})(\d+)')
#Support for overwide cellnumbering
cellnumber = re.compile('([A-Z]{1,5})(\d+)')

def wordLetter(s):
   return ord(s)-65

class SpreadsheetMap:
   def __init__(self, map, data):
      if isinstance(map, basestring):
        spreadsheet=map
        map = dict()
        map['mapping'] = dict()
        map['UseFieldNamesDirectly'] = True
        map['spreadsheet'] = spreadsheet
      self.__map = map
      self.__data = data
      self.map_expand_validate(self.__map)

   def __apply_data_to_document(self):
      filename = self.__map['spreadsheet']
      if not os.path.isfile(filename):
        raise FileNotFoundError("%s not found" % filename)
      doc = ezodf.opendoc(filename)

      if doc.doctype in ('ods', 'ots'):
         sheet = doc.sheets[0]

         for key in self.keys():
            mtype=self.gettype(key)
            if mtype in ('integer', 'number'):
               sheet[key].set_value(self[key])
            else:
               sheet[key].set_value(self[key], mtype)
      return doc

   def save(self, filename):
      self.__apply_data_to_document().saveas(filename)

   def tobytes(self):
      return self.__apply_data_to_document().tobytes()

   @staticmethod
   def mimetype():
      return 'application/vnd.oasis.opendocument.spreadsheet'

   def outputname(self):
      return self.__map['outputname']

   def __str__(self):
      return pprint.PrettyPrinter(indent=4).pformat({"map":self.__map, "data": self.__data})

   @staticmethod
   def map_from_map(entry, defaultType):
      entry.setdefault(u'combine', None)
      entry.setdefault(u'format')
      entry.setdefault(u'type', defaultType)
      entry.setdefault(u'pattern')
      multiple=entry.setdefault(u'multiple')
      if multiple:
         multiple['lut']=[
           -1 if t else None for i, t in enumerate(entry['type'])]
         j=-1
         for i,v in enumerate(multiple['lut']):
           if v:
             j+=1
             multiple['lut'][i]=j
      return entry

   def map(self):
     return copy.deepcopy(self.__map)

   @staticmethod
   def getcoord(cellname):
      fields = cellnumber.match(cellname)
      alph = fields.groups()[0]
      num = int(fields.groups()[1])

      anum = 0
      for i,c in enumerate(alph):
        anum = anum*26 + wordLetter(c) + 1
      return (anum - 1,num)

   @staticmethod
   def makecoord(coord):
      n=coord[0]
      alph=chr((n%26)+65)
      while n>=26:
         n=n//26-1
         alph=chr((n%26)+65)+alph
      return "%s%d" % (alph, coord[1])

   @classmethod
   def map_expand_validate(cls, mymap):
      mymap.setdefault('outputname', 'output.ods')
      defaultType = mymap.setdefault('defaultType', 'string')
      for m in mymap['mapping']:
         if isinstance(mymap['mapping'][m], basestring):
            cell = mymap['mapping'][m]
            mymap['mapping'][m]= {
               u'cell': mymap['mapping'][m],
            }
         
         cell = mymap['mapping'][m].setdefault(u'cell')
         if cell:
            cell = cell.upper()
            if not cellnumber.match(cell):
               raise ValueError("'%s' is not a valid cellreference" % (cell))
            mymap['mapping'][m]['cell'] = cell
         mymap[m]=cls.map_from_map(mymap['mapping'][m], mymap['defaultType'])
      return mymap

   def keys(self):
      self.genkeys()
      return self.__vkeys
   
   def genkeys(self):
      self.__rmap = {}
      for m in self.__map['mapping']:
         c = self.__map['mapping'][m]['cell']
         if c:
            self.__rmap[c]=m

         map = self.__map['mapping'][m]
         if map['multiple']:
            if map['multiple']['entry']=="row":
               ncols = len(map['type'])
               nrows = int(map['multiple']['span'])
               cellcoord = self.getcoord(map['cell'])
               for i in range (nrows):
                  for j in range (ncols):
                     if map['type'][j]:
                        mycoord = (cellcoord[0]+j, cellcoord[1]+i)
                        pcoord= self.makecoord(mycoord)
                        self.__rmap[pcoord]=m

      if self.__map['UseFieldNamesDirectly']:
         for c in self.__data.keys():
            if c not in self.__rmap and c not in self.__map['mapping'].keys():
               cu=c.upper()
               if not cellnumber.match(cu):
                  raise ValueError("'%s' is not a valid cellreference" % (c))
               if cu in self.__map['mapping']:
                  raise KeyError("%s already exists" % (c))

               if c != cu:
                  #Raise error if cu already exists
                  self.__data[cu]=self.__data[c]
                  del self.__data[c]
               entry = {
                  u'cell': cu,
               }
               t=type(self.__data[cu])
               if t == float:
                 t="number"
               elif t == int:
                 t="integer"
               else:
                 t=self.__map['defaultType']
               self.__map['mapping'][cu]=self.map_from_map(
                 entry, t)
               self.__rmap[cu]=cu

      #This is a quick and dirty solution, we could probably do this
      #less expensively at build time.
      vkeys = []
      for k in self.__rmap.keys():
         if self[k]:
            vkeys.append(k)
      self.__vkeys = vkeys

   @staticmethod
   def cast(value, mtype):
      if mtype =='integer':
         return int(value)
      elif mtype =='number':
         return float(value)
      return value

   def __getitem__(self, item):
      key = self.__rmap[item]
      return self.get_data_entry(key, item)

   def get_data_entry(self, key, item=None):
      elem = self.__map['mapping'][key]
      if not elem['multiple']:
         if not elem['combine']:
            data = self.__data[key]
         else:
            data = []
            for c in elem['combine']:
               data.append(self.get_data_entry(c))
            data = elem['format'] % tuple(data)
         data = self.cast(data, elem['type'])
         return data
      else:
         coord=self.getcoord(item)
         base=self.getcoord(elem['cell'])
         x=coord[0]-base[0]
         y=coord[1]-base[1]
         multiple=elem['multiple']
         if multiple['entry']:
            ax=multiple['lut'][x]
            ay=y
         offset=(ax,ay)
         item=None
         if y < len(self.__data[key]):
            mtype=elem['type'][x]
            item = self.cast(self.__data[key][ay][ax], mtype)
         return item

   def gettype(self, item):
      key = self.__rmap[item]
      elem = self.__map['mapping'][key]
      if not elem['multiple']:
         return elem['type']
      else:
         coord=self.getcoord(item)
         base=self.getcoord(elem['cell'])
         x=coord[0]-base[0]
         mtype=elem['type'][x]
         return mtype

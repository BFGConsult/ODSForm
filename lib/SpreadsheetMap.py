#!/usr/bin/python
# -*- coding: utf-8 -*-

import io, ezodf, sys, re, pprint
from datetime import date, datetime

try:
  basestring
except NameError:
  basestring = str

cellnumber = re.compile('([A-Z]{1,2})(\d+)')

def wordLetter(s):
   return ord(s)-65

class SpreadsheetMap:
   def __init__(self, map, data):
      self.__map = map
      self.__data = data
      self.mapExpandValidate()

   def __applyDataToDocument(self):
      doc = ezodf.opendoc(self.__map['spreadsheet'])

      if doc.doctype in ('ods', 'ots'):
         sheet = doc.sheets[0]

         for key in self.keys():
            mtype=self.getType(key)
            if mtype in ('integer'):
               sheet[key].set_value(self[key])
            else:
               sheet[key].set_value(self[key], mtype)
      return doc

   def save(self, filename):
      self.__applyDataToDocument().saveas(filename)

   def tobytes(self):
      return self.__applyDataToDocument().tobytes()

   @staticmethod
   def mimetype():
      return 'application/vnd.oasis.opendocument.spreadsheet'

   def outputname(self):
      return self.__map['outputname']

   def __str__(self):
      return pprint.PrettyPrinter(indent=4).pformat({"map":self.__map, "data": self.__data})

   def mapFromMap(self, entry):
      defaultType = self.__map['defaultType']
      entry.setdefault(u'combine', None)
      entry.setdefault(u'format')
      entry.setdefault(u'type', defaultType)
      entry.setdefault(u'pattern')
      multiple=entry.setdefault(u'multiple')
      if multiple:
         j=-1
         lut=[]
         for i in range(len(entry['type'])):
            if entry['type'][i]:
               j+=1
               lut.append(j)
            else:
               lut.append(None)
         multiple['lut']=lut
      return entry

   @staticmethod
   def getCoord(cellname):
      fields = cellnumber.match(cellname)
      alph = fields.groups()[0]
      num = int(fields.groups()[1])
      
      anum = wordLetter(alph[0])
      for i in range(1, len(alph)):
         anum = anum*26+wordLetter(alph[i])
      return (anum,num)

   @staticmethod
   def makeCoord(coord):
      n=coord[0]
      alph=chr((n%26)+65)
      while n>=26:
         n=n/26
         alph=chr((n%26-1)+65)+alph
      return "%s%d" % (alph, coord[1])

   def mapExpandValidate(self):
      self.__map.setdefault('outputname', 'output.ods')
      defaultType = self.__map.setdefault('defaultType', 'string')
      for m in self.__map['mapping']:
         if isinstance(self.__map['mapping'][m], basestring):
            cell = self.__map['mapping'][m]
            self.__map['mapping'][m]= {
               u'cell': self.__map['mapping'][m],
            }
         
         cell = self.__map['mapping'][m].setdefault(u'cell')
         if cell:
            cell = cell.upper()
            if not cellnumber.match(cell):
               raise ValueError("'%s' is not a valid cellreference" % (cell))
            self.__map['mapping'][m]['cell'] = cell
         map=self.mapFromMap(self.__map['mapping'][m])
         self.__map['mapping'][m]=map

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
               cellcoord = self.getCoord(map['cell'])
               for i in range (0, nrows):
                  for j in range (0, ncols):
                     if map['type'][j]:
                        mycoord = (cellcoord[0]+j, cellcoord[1]+i)
                        pcoord= self.makeCoord(mycoord)
                        self.__rmap[pcoord]=m

      if self.__map['UseFieldNamesDirectly']:
         for c in self.__data.keys():
            if c not in self.__rmap and c not in self.__map['mapping'].keys():
               cu=c.upper()
               if c != cu:
                  #Raise error if cu already exists
                  self.__data[cu]=self.__data[c]
                  del self.__data[c]
               print (c)
               if not cellnumber.match(cu):
                  raise ValueError("'%s' is not a valid cellreference" % (c))
               entry = {
                  u'cell': cu,
               }
               self.__map['mapping'][cu]=self.__mapFromMap(entry)
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
      return value

   def __getitem__(self, item):
      key = self.__rmap[item]
      return self.getDataEntry(key, item)

   def getDataEntry(self, key, item=None):
      elem = self.__map['mapping'][key]
      if not elem['multiple']:
         if not elem['combine']:
            data = self.__data[key]
         else:
            data = []
            for c in elem['combine']:
               data.append(self.getDataEntry(c))
            data = elem['format'] % tuple(data)
         data = self.cast(data, elem['type'])
         return data
      else:
         coord=self.getCoord(item)
         base=self.getCoord(elem['cell'])
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

   def getType(self, item):
      key = self.__rmap[item]
      elem = self.__map['mapping'][key]
      if not elem['multiple']:
         return elem['type']
      else:
         coord=self.getCoord(item)
         base=self.getCoord(elem['cell'])
         x=coord[0]-base[0]
         mtype=elem['type'][x]
         return mtype

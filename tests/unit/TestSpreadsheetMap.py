#!/usr/bin/python
import unittest, copy, json
from ODSForm import SpreadsheetMap

import pprint

SM = SpreadsheetMap.SpreadsheetMap

def genString(n, base='' , m=26, toplevel=True):
    if n==1:
        return [base+chr(i+65) for i in range(m)]
    l=genString(n-1, base, m) if toplevel else []
    for i in range(m):
        l.extend(genString(n-1, base+chr(i+65), m, False))
    return l

class TestSpreadsheetMap(unittest.TestCase):
    def setUp(self):
        self.coords=genString(2)
        with open('tests/fixtures/map.json') as map_file:
            self.mapping = json.load(map_file)
        with open('tests/fixtures/data.json') as data_file:
            self.data = json.load(data_file)

    def test_getcoords(self):
        for i,txt in enumerate(self.coords):
            coord = txt+str(1)
            c = SM.getcoord(coord)
            self.assertEqual(i, c[0], "'%s' should yield column %d" % (coord, i))

    def test_make_coords(self):
        for i,txt in enumerate(self.coords):
            coord = txt+str(1)
            c = SM.getcoord(coord)
            ncoord = SM.makecoord(c)
            self.assertEqual(coord, ncoord,
                             "'%s' should yield %s, but yielded %s" % (c, coord, ncoord ))

    def test_map_from_map(self):
        mymap = copy.deepcopy(self.mapping)
        #pprint.PrettyPrinter(indent=4).pprint(mymap)
        SM.map_expand_validate(mymap)
        #pprint.PrettyPrinter(indent=4).pprint(mymap)
        self.assertEqual(self.mapping['mapping']['address'],
                         mymap['mapping']['address']['cell'],
                         "string should be replaced with datastructure")

    def test_create(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)

    def test_tobytes(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        sm.tobytes()

            

if __name__ == '__main__':
    unittest.main()

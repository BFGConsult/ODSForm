#!/usr/bin/python
import unittest
from ODSForm import SpreadsheetMap

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

    def test_get_coords(self):
        for i,txt in enumerate(self.coords):
            coord = txt+str(1)
            c = SM.get_coord(coord)
            self.assertEqual(i, c[0], "'%s' should yield column %d" % (coord, i))

    def test_make_coords(self):
        for i,txt in enumerate(self.coords):
            coord = txt+str(1)
            c = SM.get_coord(coord)
            ncoord = SM.make_coord(c)
            self.assertEqual(coord, ncoord,
                             "'%s' should yield %s, but yielded %s" % (c, coord, ncoord ))
            

if __name__ == '__main__':
    unittest.main()

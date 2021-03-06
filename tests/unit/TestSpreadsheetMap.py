#!/usr/bin/python
# -*- coding: utf-8 -*-
import unittest, copy, json
import ezodf, os

from ODSForm.SpreadsheetMap import FileNotFoundError, SpreadsheetMap as SM

def loadOdfFromBytes(bytestring):
    dumplocal=False
    from tempfile import mkstemp
    _, tmp = mkstemp()
    with open(tmp, 'wb') as f:
        f.write(bytestring)
        f.flush()
    if dumplocal:
        with open('dump.ods', 'wb') as f:
            f.write(bytestring)
            f.flush()
    doc = ezodf.opendoc(tmp)
    os.remove(tmp)
    return doc

def genString(n, base='' , m=26, toplevel=True):
    if n==1:
        return [base+chr(i+65) for i in range(m)]
    l=genString(n-1, base, m) if toplevel else []
    for i in range(m):
        l.extend(genString(n-1, base+chr(i+65), m, False))
    return l

class TestSpreadsheetMap(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.coords = genString(2)
        with open('tests/fixtures/map.json') as map_file:
            cls.mapping = json.load(map_file)
        with open('tests/fixtures/data.json') as data_file:
            cls.data = json.load(data_file)

        mymap=dict()
        mymap['mapping']= []
        mymap['spreadsheet']="tests/fixtures/Expenselist.ots"
        cls.minimalmap=mymap


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
        SM.map_expand_validate(mymap)
        self.assertEqual(self.mapping['mapping']['address'],
                         mymap['mapping']['address']['cell'],
                         "string should be replaced with datastructure")

    def test_create(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)

    def test_nonexisting_file(self):
        mymap = copy.deepcopy(self.mapping)
        mymap['spreadsheet']='dummy.ots'
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        with self.assertRaises(FileNotFoundError):
            sm.tobytes()

    def test_invalid_cellreference(self):
        mymap = copy.deepcopy(self.mapping)
        mymap['mapping']['account']['cell']='dummy'
        mydata = copy.deepcopy(self.data)
        with self.assertRaises(ValueError):
            sm = SM(mymap, mydata)

    def checkLoadedFile(self, ods):
        sheet=ods.sheets[0]
        self.assertEqual(sheet['A4'].value,
                         'Getting drunk')
        #We should set date as well
        #print(sheet['E4'].value)
        self.assertEqual(sheet['D7'].value,
                         'John Doe')
        self.assertEqual(sheet['D8'].value,
                         '123 Main St')
        self.assertEqual(sheet['D9'].value,
                         '1234 Anytown')
        self.assertEqual(sheet['D10'].value,
                         12345678901)
        self.assertEqual(sheet['A15'].value,
                         '2018-12-24')
        self.assertEqual(sheet['C15'].value,
                         'Spirits')
        self.assertEqual(sheet['G15'].value,
                         90)
        self.assertEqual(sheet['A16'].value,
                         '2019-05-04')
        self.assertEqual(sheet['C16'].value,
                         u'Øl')
        self.assertEqual(sheet['G15'].value,
                         90)
        self.assertEqual(sheet['G16'].value,
                         123)

    def test_tobytes(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        bytes = sm.tobytes()
        ods = loadOdfFromBytes(bytes)
        self.checkLoadedFile(ods)

    def test_save(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        from tempfile import mkstemp
        _, tmp = mkstemp()
        sm.save(tmp)
        ods = ezodf.opendoc(tmp)
        os.remove(tmp)
        self.checkLoadedFile(ods)

    def test_mimetype(self):
        self.assertEqual(SM.mimetype(),
                         'application/vnd.oasis.opendocument.spreadsheet')

    def test_outname(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        self.assertEqual(sm.outputname(),
                         'Expenselist.ods')

    def test_stringrepr(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        import ast
        v=ast.literal_eval(str(sm))
        smn = SM(v['map'],v['data'])
        bytes = smn.tobytes()
        ods = loadOdfFromBytes(bytes)
        self.checkLoadedFile(ods)

    def test_map(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        smn = SM(sm.map(), mydata)
        bytes = smn.tobytes()
        ods = loadOdfFromBytes(bytes)
        self.checkLoadedFile(ods)

    def test_nomap(self):
        with open('tests/fixtures/nomap_data.json') as data_file:
            data = json.load(data_file)
        sm = SM("tests/fixtures/Expenselist.ots", data)
        bytes = sm.tobytes()
        ods = loadOdfFromBytes(bytes)
        self.checkLoadedFile(ods)

    def test_invalidcell(self):
        with open('tests/fixtures/nomap_data.json') as data_file:
            data = json.load(data_file)
        data['dummy']='test'
        sm = SM("tests/fixtures/Expenselist.ots", data)
        with self.assertRaises(ValueError):
            bytes = sm.tobytes()

    def test_duplicatecell(self):
        with open('tests/fixtures/nomap_data.json') as data_file:
            data = json.load(data_file)
        data['a4']='duplicate'
        with self.assertRaises(KeyError):
            sm = SM("tests/fixtures/Expenselist.ots", data)

    def test_nonexistant_datapoint(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        mymap['mapping']['dummy']='E4'
        sm = SM(mymap, mydata)
        try:
            with self.assertWarns(Warning):
                bytes = sm.tobytes()
        except:
            raise unittest.SkipTest('assertWarns not supported')

    def test_nonexistant_mappoint(self):
        mymap = copy.deepcopy(self.mapping)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        with self.assertRaises(KeyError):
            foo = sm['dummy']

    def test_minimal(self):
        mymap = copy.deepcopy(self.minimalmap)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        bytes = sm.tobytes()

    def test_subminimal(self):
        mymap = copy.deepcopy(self.minimalmap)
        mydata = copy.deepcopy(self.data)
        sm = SM(mymap, mydata)
        bytes = sm.tobytes()

    def test_subminimal1(self):
        mymap = copy.deepcopy(self.minimalmap)
        mydata = copy.deepcopy(self.data)
        del mymap['spreadsheet']
        with self.assertRaises(ValueError):
            sm = SM(mymap, mydata)

    def test_subminimal2(self):
        mymap = copy.deepcopy(self.minimalmap)
        mydata = copy.deepcopy(self.data)
        del mymap['mapping']
        with self.assertRaises(ValueError):
            sm = SM(mymap, mydata)

if __name__ == '__main__':
    unittest.main()

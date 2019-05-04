#!flask/bin/python
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

import json, sys
from flask import Flask, request, make_response, render_template
from SpreadsheetMap import SpreadsheetMap

app = Flask(__name__)

with open(sys.argv[3]) as map_file:
    mapping = json.load(map_file)
mapping=SpreadsheetMap.mapExpandValidate(mapping)
app = Flask(__name__)

@app.route('/', methods=['POST', 'GET'])
def post():
    if request.method == 'POST':
        if not request.is_json:
            app.logger.info('Request is not json', user.username)
        else:
            content = request.get_json()

            print ("Generating report for '%s'" %( content['purpose']))
            print (content)

            smap = SpreadsheetMap(mapping, content)

            document = smap.tobytes()
            response = make_response(document)
            response.headers.set('Content-Type', smap.mimetype())
            response.headers.set(
                'Content-Disposition', 'attachment', filename=smap.outputname())
            return response
    else:
        res='<table><form>'
        for k,v in mapping['mapping'].items():
            if not v.get('combine', None):
                if v.get('multiple', None):
                    res+='<tr><th>%s</th></tr>'%(k)
                    for i in range(v['multiple']['span']):
                        res+='<tr>'
                        for t,l in zip(v['type'], v['label']):
                            if t:
                                res+='<td>%s</td><td><input></input></td>'%(l)
                        res+='</tr>'
                else:
                    res+='<tr><td>%s</td><td><input></input></td></tr>'%(k)
        res+='</form></table>'
        return make_response(res, 200)

app.run(host=sys.argv[1], port=int(sys.argv[2]))

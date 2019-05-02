#!flask/bin/python
from flask import Flask
import ezodf
import json
from SpreadsheetMap import SpreadsheetMap

app = Flask(__name__)

from flask import Flask
from flask import request, make_response
import sys

with open(sys.argv[3]) as map_file:
    mapping = json.load(map_file)

app = Flask(__name__)

@app.route('/fillform', methods=['POST'])
def post():
    if not request.is_json:
        app.logger.info('Request is not json', user.username)        
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

app.run(host=sys.argv[1], port=int(sys.argv[2]))

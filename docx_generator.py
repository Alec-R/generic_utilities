# Generate a filled docx from a template based on api or database output

import requests
import json
import dataclasses

from docxtpl import DocxTemplate  



def from_template(template, name, payload):
    """
    render docx from template and payload
    template: full path to template file
    name: full path to output file
    payload: dict of variables
    """
    name = name + '.docx'
    fullpath = name
    template = DocxTemplate(template)
    template.render(payload, autoescape=True)
    template.save(fullpath)
    print("document saved at ", fullpath)

def save_json_payload(jason, filename, folderpath):
    # file names for output
    jsonfilename = f"{filename}.json"
    jsonpath = folderpath + jsonfilename
    with open(jsonpath, 'w') as f:
        json.dump(jason, f)

def docxDocs(prod):
    """generates TS and FT for a product, save the product structure"""
    # TODO: amend prod.template to a generic path variable name
    from_template(prod.FTtemplate, prod.finaltermspath, prod.docpayload)
    from_template(prod.TStemplate, prod.termsheetpath, prod.docpayload)

    jason = dataclasses.asdict(prod)
    filename = f"Product structure {prod.rj['series']}"
    save_json_payload(jason, filename, prod.folderpath)

def fetch_from_api(identifier):
    """
    read database and return JSON output
    """
    url = 'your_url'

    full_url = url +'/'+ identifier

    r = requests.get(full_url)
    rj = r.json()
    return rj

def fetch_from_db(identifier):
    """ read SQL"""
    return None
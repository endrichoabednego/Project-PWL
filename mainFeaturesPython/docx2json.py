import json
import zipfile
from lxml import etree
import docx
import argparse
import os
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import xml.etree.ElementTree as ET
from docx2python import docx2python
import io

# create argument parser
parser = argparse.ArgumentParser(description='Convert a docx file to a json file.')
parser.add_argument('-i', '--input', type=str, required=True, help='input docx file path')
parser.add_argument('-o', '--output', type=str, required=True, help='output json file path')
args = parser.parse_args()

# extract input and output file paths from arguments
input_file = args.input
output_file = args.output
# Open the document using the python-docx library
doc =docx.Document(input_file)

# get the first paragraph in the document
paragraph = doc.paragraphs[0]

# get the first run in the paragraph
run = paragraph.runs[0]

# get the properties of the run
run_props = run._element.rPr

# create a new w:noProof element
no_proof = OxmlElement('w:noProof')

# set the w:val attribute to true
no_proof.set(qn('w:val'), 'true')

# add the w:noProof element to the run's properties
run_props = docx.oxml.shared.OxmlElement('w:rPr')
no_proof = docx.oxml.shared.OxmlElement('w:noProof')
no_proof.set(docx.oxml.ns.qn('w:val'), 'true')
run_props.append(no_proof)

# save the modified document




if doc.save:
    print(input_file)
        # create argument parser
    parser = argparse.ArgumentParser(description='Convert a docx file to a json file.')
    parser.add_argument('-i', '--input', type=str, required=True, help='input docx file path')
    parser.add_argument('-o', '--output', type=str, required=True, help='output json file path')
    args = parser.parse_args()

    # extract input and output file paths from arguments
    input_file = args.input
    output_file = args.output

    zip_file = zipfile.ZipFile(input_file)

    # membaca file XML dari dalam zip file
    xml_file = zip_file.read('word/document.xml')
    xml_doc = etree.fromstring(xml_file)

    # menentukan namespace yang digunakan dalam file XML
    namespace = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # melakukan ekstraksi menggunakan XPath
    xquery_expr = "//w:p[@w:rsidR]"
    result = xml_doc.xpath(xquery_expr, namespaces=namespace)

    # menyimpan hasil ekstraksi ke dalam JSON
    sentences = []
    prev_rsidR = ''
    counter=0
    for elem in result:
        text = ''
        for child in elem.iter():
            if child.tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t':
                rsidR = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rsidR')
                if rsidR != prev_rsidR:
                    if text:
                        sentences.append(text)
                        text = ''
                    prev_rsidR = rsidR
                if text:
                    if child.getparent().tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r':
                        if child.text!=" ":
                            text += f"<span id='{counter}'>"
                            # if child.getparent().find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr') is not None:
                            #     rpr = child.getparent().find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                            #     if rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b') is not None:
                            #         text += f'<strong>{child.text}</strong>'
                            #     elif rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i') is not None:
                            #         text += f'<em>{child.text}</em>'
                            #     elif rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}u') is not None:
                            #         text += f'<u>{child.text}</u>'
                            #     else:
                            #         text += child.text
                            text += child.text
                            text +="</span>"
                            counter+=1
                        else:
                            text+=child.text
                            counter+=1
                    else:
                        text += child.text
                        counter+=1

                        
                else:
                    if child.getparent().tag == '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r':
                        # if child.getparent().find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr') is not None:
                        #     rpr = child.getparent().find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
                        #     if rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}b') is not None:
                        #         text += f"<strong>{child.text}</strong>"
                        #     elif rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}i') is not None:
                        #         text += f"<em>{child.text}</em>"
                        #     elif rpr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}u') is not None:
                        #         text += f"<u>{child.text}</u>"
                        #     else:
                        #         text += child.text
                        # else:
                        text += child.text
                        counter+=1
                    else:
                        text += child.text
                        counter+=1
        if text:
            sentences.append(text)

    json_data = {
        "id": "",
        "timestamp": 1681489326,
        "extractedFrom": output_file,
        "sentences": sentences
    }

    # menyimpan hasil ekstraksi ke dalam file JSON
    with io.open(output_file, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)

    # menutup file zip
    zip_file.close()

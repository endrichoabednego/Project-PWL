from flask import Flask, request, jsonify
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

app = Flask(__name__)

@app.route('/convert', methods=['POST'])
def convert_docx_to_json():
    # Access the uploaded file from the request
    docx_file = request.files['file']

    # Save the uploaded file to a temporary location
    temp_file_path = 'temp.docx'
    docx_file.save(temp_file_path)

    # Define the output file path
    output_file_path = 'output.json'

    # Open the document using the python-docx library
    doc = docx.Document(temp_file_path)

    # get the first paragraph in the document
    paragraph = doc.paragraphs[0]

    # get the first run in the paragraph
    run = paragraph.runs[0]

    # get the properties of the run
    run_props = run._element.rPr

# Check if rPr element exists before appending
    if run_props is None:
    # Create rPr element if it doesn't exist
        run._element.rPr = OxmlElement('w:rPr')
        run_props = run._element.rPr

# create a new w:noProof element
    no_proof = OxmlElement('w:noProof')
    no_proof.set(qn('w:val'), 'true')

# add the w:noProof element to the run's properties
    run_props.append(no_proof)


    # save the modified document
    doc.save(temp_file_path)

    # Perform the conversion using your existing code
    # Replace the existing print statements with appropriate code to save the JSON output

    # Perform the conversion

    zip_file = zipfile.ZipFile(temp_file_path)

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
        "extractedFrom": output_file_path,
        "sentences": sentences
    }

    # save the JSON output to the specified output_file_path
    with io.open(output_file_path, 'w', encoding='utf-8') as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)

    # Close the zip file
    zip_file.close()

    # Return the JSON response
    return jsonify({'message': 'Conversion successful'})

if __name__ == '__main__':
    
    app.run()

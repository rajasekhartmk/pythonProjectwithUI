import os

import en_core_web_sm

nlp = en_core_web_sm.load()
from flask_restful import Resource, Api, reqparse
import threading
import subprocess
import sys
from flask import Flask, render_template
import Resume_service
import ConvertToPdf

controller = Flask(__name__)
api = Api(controller)

@controller.route('/')
def index():
    return render_template('index.html')

# @controller.route('/convert/')
# def convert():
#     def run_script():
#         theproc = subprocess.Popen([sys.executable, "pdfconverter.py"])
#         theproc.communicate()
#
#     threading.Thread(target=lambda: run_script()).start()
#     return render_template('convert.html')

api.add_resource(ConvertToPdf.ConvertToPdf, '/data/')
api.add_resource(Resume_service.ResumeScreen, '/check/')

if __name__ == "__main__":
    controller.run(
        host="127.0.0.1",
        port=int("5000"),
        debug=True
    )
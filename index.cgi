#!/usr/bin/env python

import csv
import operator
import subprocess
import sys
import time
from wsgiref.handlers import CGIHandler

sys.path.extend([
    'click-6.7',
    'Flask-0.12',
    'itsdangerous-0.24',
    'Werkzeug-0.12.1',
])

from flask import Flask
from flask import redirect
from flask import render_template
from flask import request

app = Flask(__name__)


@app.template_filter()
def support_level(value):
  if value.endswith(' Support'):
    value = value[:-len(' Support')]
  return value


@app.template_filter()
def contact_information(data):
  value = data['Home Phone']
  if value:
    yield 'H:' + value
  value = data['Mobile Phone']
  if value:
    yield 'C:' + value
  value = data['Email Address']
  if value:
    yield value


@app.route('/', methods=['GET', 'POST'])
def index():
  if request.method == 'POST':
    stem = time.strftime('%Y%m%d_%H%M%S', time.gmtime())
    filename = stem + '.csv'
    request.files['file'].save(filename)
    with open(filename) as f:
      reader = csv.DictReader(f)
      filename = stem + '.html'
      open(filename, 'w').write(
          render_template(
              'result.html',
              reader=sorted(
                  reader,
                  key=operator.itemgetter('Street Name', 'House Number',
                                          'House Unit', 'Surname',
                                          'First Name'))))
    args = ['prince-11.1-linux-generic-x86_64/lib/prince/bin/prince', filename]
    subprocess.Popen(args)
    filename = stem + '.pdf'
    return redirect('../' + filename)
  return render_template('index.html')


CGIHandler().run(app)

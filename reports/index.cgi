#!/usr/bin/env python

import ConfigParser
import string
import sys
from wsgiref.handlers import CGIHandler

sys.path[0:0] = [
    'click-6.7',
    'Flask-0.12',
    'google-api-python-client-1.6.2',
    'itsdangerous-0.24',
    'Jinja2-2.9.6',
    'oauth2client-4.0.0',
    'six-1.10.0',
    'uritemplate-3.0.0',
    'Werkzeug-0.12.1',
]

from apiclient import discovery
import httplib2
from flask import Flask
from flask import g
from flask import render_template
from oauth2client.file import Storage

app = Flask(__name__)
title = 'Data'


def get_spreadsheet_id():
  spreadsheet_id = getattr(g, 'spreadsheet_id', None)
  if spreadsheet_id is None:
    config = ConfigParser.ConfigParser()
    config.read('config.ini')
    g.spreadsheet_id = spreadsheet_id = config.get('DEFAULT', 'spreadsheet_id')
  return spreadsheet_id


def get_service():
  service = getattr(g, 'service', None)
  if service is None:
    filename = 'credentials'
    storage = Storage(filename)
    credentials = storage.get()
    http = credentials.authorize(httplib2.Http())
    g.service = service = discovery.build('sheets', 'v4', http=http)
  return service


def grid_range(start_row, start_column, end_row, end_column):
  return '{title}!{start_column}{start_row}:{end_column}{end_row}'.format(
      title=title,
      start_column=string.ascii_uppercase[start_column],
      start_row=start_row + 1,
      end_column=string.ascii_uppercase[end_column - 1],
      end_row=end_row)


def read():
  result = get_service().spreadsheets().get(
      spreadsheetId=get_spreadsheet_id()).execute()
  for sheet in result['sheets']:
    properties = sheet['properties']
    if properties['title'] == title:
      break
  grid_properties = properties['gridProperties']
  row_count = grid_properties['rowCount']
  column_count = grid_properties['columnCount']
  result = get_service().spreadsheets().values().get(
      spreadsheetId=get_spreadsheet_id(),
      range=grid_range(0, 0, row_count, column_count)).execute()
  return result['values']


def get_rows():
  result = read()
  fieldnames = result[0]
  for row_data in result[1:]:
    row_data = dict(zip(fieldnames, row_data))
    if row_data.get('Precinct') and row_data.get('Support'):
      value = row_data.get('2017_gotv_voted')
      if value:
        value = set(value.split(', '))
        if not value.isdisjoint([
            '20170429',
            '20170430',
            '20170503',
            '20170504',
            '20170505',
            '20170506',
        ]):
          row_data['2017_gotv_voted'] = 'Advanced'
          yield row_data
        elif 'Voted' in value:
          row_data['2017_gotv_voted'] = 'General'
          yield row_data


@app.route('/')
def index():
  return render_template(
      'index.html',
      rows=list(get_rows()),
      support_levels=[
          'Support Strong',
          'Support Weak',
          'Undecided',
          'Opposed Weak',
          'Opposed Strong',
      ])


CGIHandler().run(app)

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
    'oauth2client-4.0.0',
    'six-1.10.0',
    'uritemplate-3.0.0',
    'Werkzeug-0.12.1',
]

from apiclient import discovery
import httplib2
from flask import Flask
from flask import g
from flask import redirect
from flask import render_template
from flask import request
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


def grid_coordinate(row, column):
  return '{title}!{column}{row}'.format(
      title=title, column=string.ascii_uppercase[column], row=row + 1)


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


def sort_key(value):
  i, row_data = value
  return row_data['2017_ebc_vsn'], row_data['LAST NAME'], row_data['FIRST NAME']


@app.template_filter()
def voted(value):
  return not {
      '20170429',
      '20170430',
      '20170503',
      '20170504',
      '20170505',
      '20170506',
      'Voted',
  }.isdisjoint(value.split(', '))


@app.template_filter()
def street_address(row_data):
  return '-'.join(value
                  for value in [
                      row_data.get('Unit'),
                      ' '.join(value for value in [
                          row_data.get('House Number'),
                          row_data.get('Street Name'),
                      ] if value),
                  ] if value)


@app.route('/')
def index():
  result = read()
  fieldnames = result[0]
  j = fieldnames.index('Precinct')
  return render_template(
      'index.html',
      precincts=sorted({
          row_data[j]
          for row_data in result[1:] if len(row_data) > j and row_data[j]
      }))


@app.route('/<precinct>', methods=['GET', 'POST'])
def precinct(precinct):
  result = read()
  fieldnames = result[0]
  rows = [(i, dict(zip(fieldnames, row_data)))
          for i, row_data in enumerate(result[1:], 1)]
  rows = [(i, row_data) for i, row_data in rows
          if row_data.get('Precinct') == precinct]
  for i, row_data in rows:
    try:
      row_data['2017_ebc_vsn'] = int(row_data['2017_ebc_vsn'])
    except ValueError:
      pass
  rows.sort(key=sort_key)

  if request.method == 'POST':
    if any(
        form_value != format(i)
        for form_value, (i, row_data) in zip(request.form.getlist('i'), rows)):
      return 'FIXME'
    data = []
    j = fieldnames.index('2017_gotv_voted')
    for i, row_data in rows:
      form_value = request.form.get(format(i), '')
      if form_value != row_data['2017_gotv_voted']:
        data.append(dict(range=grid_coordinate(i, j), values=[[form_value]]))
    if data:
      result = get_service().spreadsheets().values().batchUpdate(
          spreadsheetId=get_spreadsheet_id(),
          body=dict(
              valueInputOption='USER_ENTERED', data=data)).execute()
    return redirect(precinct)

  return render_template('precinct.html', rows=rows)


CGIHandler().run(app)

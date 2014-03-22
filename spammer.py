#!/usr/bin/env python3

##########################################################################
#                                                                        #
# The MIT License (MIT)                                                  #
#                                                                        #
# Copyright (c) 2014 Foivos S. Zakkak <foivos@zakkak.net>                #
#                                                                        #
# Permission is hereby granted, free of charge, to any person            #
# obtaining a copy of this software and associated documentation files   #
# (the "Software"), to deal in the Software without restriction,         #
# including without limitation the rights to use, copy, modify, merge,   #
# publish, distribute, sublicense, and/or sell copies of the Software,   #
# and to permit persons to whom the Software is furnished to do so,      #
# subject to the following conditions:                                   #
#                                                                        #
# The above copyright notice and this permission notice shall be         #
# included in all copies or substantial portions of the Software.        #
#                                                                        #
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,        #
# EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF     #
# MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND                  #
# NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE #
# LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION #
# OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION  #
# WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.        #
##########################################################################

##
# Spammer
#
# @author Foivos S. Zakkak <foivos@zakkak.net>
#
# This python script parses excel files and sends an e-mail with the
# grades to each student.  The scripts excepts to find one student per
# row.
#
# TODOs:
#    1. Add support for OpenOffice ODS

import argparse
import collections
import smtplib
import re

from email.mime.text import MIMEText
from email import charset
from time import sleep
from xlrd import open_workbook,xlsx

smtp_server   = 'mailserver.example.com'
lesson_prefix = 'cs'

##
# Ask a yes/no question via input() and return the user's answer.
# Inspired by http://code.activestate.com/recipes/577058/
#
# @param msg The question printed to the user.
# @param default The default answer if the user just hits <Enter>.
#     Can be "yes" (the default), "no" or None (meaning an answer
#     is required of the user).
#
# @return Returns True if the user was affirmative or False if she was
#         negative.
def yes_or_no(question, default="yes"):

    valid = {"yes":True,   "y":True,  "yup":True,
             "no":False,   "n":False, "nope":False}
    if default == None:
        prompt = " [y/n] "
    elif default == "yes":
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/N] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        print(question + prompt)
        choice = input().lower()
        if default is not None and choice == '':
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            print("Please respond with 'yes' or 'no' "\
                "(or 'y' or 'n').\n")

##
# Parse the assignment column, which are in the form A:F, :AB, CA: or DF
#
# @param assignment_column  The string to parse.
# @param email_column       The column containing the e-mails.
# @param sheet_columns      The number of total columns in the sheet.
#
# @return Returns a list of integers with all the columns that should be parsed
def parse_assigment_column(assignment_column, email_column, sheet_columns):
    if assignment_column.isalpha():
        zero,col = xlsx.cell_name_to_rowx_colx(assignment_column+'1')
        assert(zero == 0)
        return [col]

    ranges = assignment_column.split(':',1)

    if len(ranges) != 2:
        raise Exception('Failed to parse assignment-columns'
                        ' (%s)' % assignment_column)

    start = ranges[0]
    end   = ranges[1]

    if start.isalpha():
        zero,start = xlsx.cell_name_to_rowx_colx(start+'1')
        assert(zero == 0)
    elif start == '':
        start = email_column
    else:
        raise Exception('Failed to parse assignment-columns'
                        ' (%s)' % assignment_column)

    if end.isalpha():
        zero,end = xlsx.cell_name_to_rowx_colx(end+'1')
        assert(zero == 0)
    elif end == '':
        end = sheet_columns
    else:
        raise Exception('Failed to parse assignment-columns'
                        ' (%s)' % assignment_column)

    if start < end :
        raise Exception('In ranges (:) the first column must be smaller than'
                        ' the second (in alphabetical order). The offending'
                        ' argument is \"%s\"' % assignment_column)

    return [range(start, end+1)]

##
# Parse the assignment columns, which are in the form A:F,:AB,CA:,DF etc.
#
# @param assignment_columns The string to parse.
# @param email_column       The column containing the e-mails.
# @param sheet_columns      The number of total columns in the sheet.
#
# @return Returns a list of integers with all the columns that should be parsed
def parse_assigment_columns(assignment_columns, email_column, sheet_columns):

    # if there where no columns specified for assignments' grades
    if assignment_columns == []:
        return []

    recurse = parse_assigment_columns(assignment_columns[1:],
                                      email_column,
                                      sheet_columns)
    result  = parse_assigment_column(assignment_columns[0],
                                     email_column,
                                     sheet_columns)
    result.append(recurse)
    return result

##
# Validates an email address
#
# @param email The string to be validated
#
# @return True if it is validated, False otherwise
def is_valid_email(email):

    email_re = re.compile('[^@]+\@[^@]+\.[^@][^@]+')

    if len(email) > 6:
        if not email_re.match(email) is None:
            return True

    return False

##
# Takes a list with nested lists and flattens it (makes it a list of items)
#
# @param l The list to be flattened
#
# @return The flattened list (as an iterable collection)
def flatten(l):
    for el in l:
        if isinstance(el, collections.Iterable):
            for sub in flatten(el):
                yield sub
        else:
            yield el

################################################################################
# Define the arguments
################################################################################
parser = argparse.ArgumentParser(prog="spammer.py", description='The Spammer!!!')
parser.add_argument('spreadsheet',
                    type=str)
parser.add_argument('-s','--sheet',
                    type=int,
                    nargs=1,
                    default=0,
                    help='choose the sheet to parse in range 0..99 (default: 0)')
parser.add_argument('-H','--header-row',
                    type=int,
                    required=True,
                    help='choose the header row, assignment names will be read '
                    'from there.  All rows bellow it will be parsed.'
                    '(default: 0)')
parser.add_argument('-e','--email-column',
                    type=str,
                    required=True,
                    help='choose the column containing the students\' e-mails.')
parser.add_argument('-l','--lesson',
                    type=int,
                    required=True,
                    help='the lesson number (i.e. 255).')
parser.add_argument('-c','--assignment-columns',
                    type=str,
                    help='choose the columns containing the assignments\' grades.'
                    ' Supports comma separated values (i.e. A,C,D) and ranges '
                    '(i.e. A,C:D). '
                    '(default: Will parse all columns after the e-mail column)')
parser.add_argument('-D','--dry-run',
                    dest='dry',
                    action='store_true',
                    help='perform a dry run (default: False)')
parser.add_argument('-f','--force',
                    action='store_true',
                    help='force the execution without asking for confirmation.'
                    '(default: False)')
parser.add_argument('-v','--verbose',
                    dest='verbose',
                    action='store_true',
                    help='run in verbose mode (default: False)')
parser.add_argument('-V', '--version',
                    action='version',
                    version='Spammer v0.9')

args = parser.parse_args()

if args.verbose:
    print('Args = ' + str(args))

zero,email_column = xlsx.cell_name_to_rowx_colx(args.email_column+'1')
assert(zero == 0)

if args.assignment_columns is None:
    args.assignment_columns = ':'

header_row = args.header_row - 1

# open the spreadsheet
book  = open_workbook(filename=args.spreadsheet)
# go to the proper sheet
sheet = book.sheet_by_index(args.sheet)

if not args.force:
    print('You chose to open sheet (' + str(args.sheet) + ') "' +\
        sheet.name + '".')
    yes_or_no('Are you sure you want to proceed with this choice?')

assignments = parse_assigment_columns(args.assignment_columns.split(','),
                                      email_column,
                                      sheet.ncols)

assignments = list(flatten(assignments))

if args.verbose:
    print('Will send grades for: ', assignments)

sender      = lesson_prefix.lower() + str(args.lesson) + '@example.com'
subject     = lesson_prefix.upper() + '-' + str(args.lesson) + ' Βαθμολογίες ( '
# add info about which assignments' grades are sent
for col in assignments:
    subject = subject + str(sheet.cell(header_row, col).value) + ' '
subject     = subject + ')'

charset.add_charset('utf-8', charset.BASE64, charset.BASE64)

if args.verbose:
    print('Subject: ',subject)

# for all rows bellow the header
for row in range(header_row + 1, sheet.nrows):
    to = sheet.cell(row, email_column).value

    if not is_valid_email(to):
        raise Exception('\"' + to  + '\" is not a valid e-mail address')

    text = 'Καλησπέρα,\n\n'
    text = text + 'Ακολουθούν οι βαθμολογίες:\n\n'

    # for all assignments
    for col in assignments:
        text = text + str(sheet.cell(header_row, col).value)
        text = text + '\t' + str(sheet.cell(row, col).value) + '\n'

    msg = MIMEText(text, _charset='utf-8')
    msg['Subject'] = subject
    msg['From']    = sender
    msg['To']      = to

    if args.verbose:
        print('Sending to: ',to)
#        print(msg.as_string())

    if not args.dry:
        s = smtplib.SMTP(smtp_server)
        s.sendmail(sender, [to], msg.as_string())
        s.quit()
        # wait a sec before sending the next e-mail (to avoid
        # overloading the mail server)
        sleep(1)

# grade_spammer

A python script parsing excel files with student grades and sending
out the grades with personal e-mails

## Dependencies

1. python3
2. python3-xlrd

## Usage

```
usage: spammer.py [-h] [-s SHEET] -H HEADER_ROW -e EMAIL_COLUMN -l LESSON
                  [-c ASSIGNMENT_COLUMNS] [-D] [-f] [-v] [-V]
                  spreadsheet

The Spammer!!!

positional arguments:
  spreadsheet

optional arguments:
  -h, --help            show this help message and exit
  -s SHEET, --sheet SHEET
                        choose the sheet to parse in range 0..99 (default: 0)
  -H HEADER_ROW, --header-row HEADER_ROW
                        choose the header row, assignment names will be read
                        from there. All rows bellow it will be
                        parsed.(default: 0)
  -e EMAIL_COLUMN, --email-column EMAIL_COLUMN
                        choose the column containing the students' e-mails.
  -l LESSON, --lesson LESSON
                        the lesson number (i.e. 255).
  -c ASSIGNMENT_COLUMNS, --assignment-columns ASSIGNMENT_COLUMNS
                        choose the columns containing the assignments' grades.
                        Supports comma separated values (i.e. A,C,D) and
                        ranges (i.e. A,C:D). (default: Will parse all columns
                        after the e-mail column)
  -D, --dry-run         perform a dry run (default: False)
  -f, --force           force the execution without asking for
                        confirmation.(default: False)
  -v, --verbose         run in verbose mode (default: False)
  -V, --version         show program's version number and exit
```

## Examples

Assume a `test.xls` file with the following contents:

|   | A       | B            | C    | D                 | E              | F              | D              |
|---|---------|--------------|------|-------------------|----------------|----------------|----------------|
| 1 | **Name** | **Surname** | **ID** | **e-mail** | **Assignment 1** | **Assignment 2** | **Assignment 3** |
| 2 | John    | Smith        | 1600 | john@example.com  | 4              | 8              | 9              |
| 3 | Μιράντα | Παπαδοπούλου |  524 | mpapa@example.com | 6              | 7.5            | 9.9            |
| 4 | Foivos  | Zakkak       |  642 | foivos@zakkak.net | 3.5            | 4              | 6              |

To send the grades for assignments 1 and 2 you can run:

```
spammer.py test.xls -H 1 -e D -l 255 -c E:F
```

or

```
spammer.py test.xls -H 1 -e D -l 255 -c :F
```

in the second case the script will print all columns after the one containig the e-mail until column F.

To send the grades for assignments 1 and 3 you can run:

```
spammer.py test.xls -H 1 -e D -l 255 -c E,D
```

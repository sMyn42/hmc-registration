# HMC Registration Tool

### Purpose

The purpose of the HMC Registration Tool is to take a series of school-separated role request forms and create committee assignments.

### Data Format

The data should be school separated by school and consists of spreadsheets with the following columns: School Name, Delegate First Name, Delegate Last Name, Delegate Grade, HMC Experience, Delegate Email, First Committee Preference, Second Committee Preference, Third Committee Preference, Fourth Committee Preference, Fifth Committee Preference, Sixth Committee Preference, Seventh Committee Preference, Eighth Committee Preference.

The files should be in .csv or .xlsx formats.

### Usage

First, run ```pip install requirements.txt``` with python3 in the base directory.

Simply ensure that the ```main.py``` script is in the base directory. Ensure that new role requests get placed in the ```pending-role-requests``` folder.

Then, run the script in the command line. A new spreadsheet should be created in the results folder with desired placements.

Subsequent requests forms should be downloaded to the ```pending-role-requests``` folder and the script should be re-run.

If the whole request processing program needs to be run from scratch, move all sheets to the ```pending-role-requests``` folder, delete the ```committee_assignments.xlsx``` file, and run the program.

Note that some committees may not end up full enough to be viable; the students in such committees must be manually redistributed to others by the user.

The ```caps``` dictionary contains caps for each committee; these can be altered as needed.

### Relevant Files

```utils.py```
```create_sheet.py```
```params.py```
```process_requests.py```
```requirements.txt```
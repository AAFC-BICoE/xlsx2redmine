# XLSX2Redmine

XLSX2Redmine automates importing tasks into Redmine by parsing a spreadsheet file in XLSX format and using Redmine's webservices to create the issues.  This tool permits importing a work breakdown structure exported from MS Project as XLSX.

Features:
* Assign tracker
* Set assignee
* Set predecessors
* Set start and due dates
* Set parent tasks


Limitations:
* Tasks can only be imported into one project at a time
* Only predecessor relationships are supported
* Only a default tracker is supported

## Installation

The following installation procedures require git, Python 2.7, virtualenv, and pip.  Please ensure these are installed prior to proceeding.

```bash
git clone https://github.com/AAFC-MBB/xlsx2redmine
cd xlsx2redmine
virtualenv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Usage

### Requirements:

* Python 2.7 (Other versions may work but are not tested) 
* A redmine deployment (should work with 2.6+ but has only been tested on 2.6)
* A Redmine administrator account
* Redmine webservices API must be enabled from the administrator interface

### Step by step:

* Export a work breakdown structure from MS Project.  Ensure that you have the following fields in the excel sheet.
** Identifier: Unique integer value that identifies each task (e.g.: 1)
** WBS number: A period seperated string value that indicates parent-child relationships. (e.g. A.b.a and A.b.b are children of A.b).  Alpha-numerical string values are supported (e.g. 1.A.2) but the seperator must be a period.
** Subject: A string that will map to the issue title.
** Start Date: A date formatted field
** Due Date: A date formatted field
** Assignee: A "firstname lastname" formatted string that maps to the name of a user in redmine.
** Predecessors: A comma separated list of identifiers for preceding tasks (e.g.: 1,2).  These must refer to the values in the Identifier column.

	NOTE: You may need to manipulate dates exported from MS Project in order for excel to recognize them as a date field.

* Create a project in redmine and take note of the project identifier

* Add members to the project that will be assigned tasks (will no longer be required in the future).  
	NOTE: Issues cannot be be assigned to non-members.  If an assignee is not found or the first name and last name do not match, a warning message will be printed and an assignee will not be set.

* Copy the config.yml.sample file to config.yml and edit it to fill in all the fields
** Set the spreadsheet path and the sheet name to use within the workbook
** Ensure that the Redmine URL is populated and either a username/password or API key must be provided.
** You must map the fields under the "mapping" section based on column names. (e.g. For the identifier, it is located in column A.)
	WARNING:  Username and Password authentication with the LDAP module may cause the directory to lock the account due to too many authentication requests within a short period of time.  It is recommended to use an API key as an alternative or to use a local redmine account.
** Set the project identifier, default tracker

* Execute xslx2redmine.py and pass it the configuration file as a parameter

	python xslx2redmine.py -c config.yml | tee xlsx2redmine.log

## ToDo

* Automatically create projects that don't exist
* Automatically add members to the project
* Multiple tracker support with a default tracker

## Contributing
1. Fork it!
2. Create your feature branch: `git checkout -b my-new-feature`
3. Commit your changes: `git commit -am 'Add some feature'`
4. Push to the branch: `git push origin my-new-feature`
5. Submit a pull request



## Credits

Author(s): 
* Iyad Kandalaft <iyad.kandalaft@canada.ca>

Contriutor(s):
* This work would not be possible without the author of the [Python Redmine](https://github.com/maxtepkeev/python-redmine) project

## License

The MIT License (MIT)

Copyright (c) 2016 Government of Canada

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.



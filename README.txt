===============================================================================
**************************AUTOMATED GEOCODER***********************************
===============================================================================

Last Update: 02/23/2018
Author: Jonathan Engelbert (Jonathan.Engelbert@sfgov.org)

**********************************************************

FUNCTIONALITY:

1 -> Parses through Excel spreadsheet looking for address data

	* Column must be named "Address", "address", "Location" or "location"
	* Target column name above must be unique
	* Excel file must be name "test.xlsx"
	* .xls files are not accepted

2 -> Standardizes address input observing the Enterprise Address System database format

 * Address types are abbreviated, following the standards found here:
   
 https://data.sfgov.org/Geographic-Locations-and-Boundaries/Street-Names/6d9h-4u5v/data

 * Everything to the right of an address type is wiped out (as in
     "123 1st street, Apt 345"  --> "123 1st st"
 * Apostrophes are removed and leave no whitespace in between characters
 * Ranges (as in "517-520 1st Street") are eliminated. Only the first
   number in the range must is kept
 * No fractions in addresses (as in 10/2 Market Street)
 * Streets and avenues from 1-9 must have leading zeroes (as in
   7th street --> 07th street
 * Cardinal directions are kept to the right of the address type in 13
   cases
 * Abbreviates addressees and wipes out characters to the right of
     address type
 * Makes sure that cardinal directions are kept (as in 123 1st NE, or 1 S     
   Van Ness Ave)
 * Remove apostrophes, periods, octothorpe and commas
 * Removes ranges
 * Removes letters mixed with numbers
 * Removes fractions in addresses
 * Adds a number to single digit addresses (as in 7th ave * 07th Ave)
 * Transforms BAYSHORE BLVD --> BAY SHORE BLVD
 * Handles Embarcadero Center Transformations as:
      1 embarcadero center --> 301 CLAY ST
      2 embarcadero center –-> 201 CLAY ST
      3 embarcadero center –-> 101 CLAY ST
      4 embarcadero center –-> 150 DRUMM ST
 * Handles addresses commonly entered without address type
 * Handles assorted edge cases as they are identified

3 -> Generates new spreadsheet 

* Writes new column to the next available column in original spreadsheet
* Saves workbook changes and writes address Transformations
* New spreadsheet generated is named "transformed.xlsx"

4 -> Geocodes spreadsheet (optional)

* Program looks for excel file named "transformed.xls"
* Converts to table
* Geocodes using EAS geocoder 
* Records that pass geocoder test:
  - Get an added field called "geocoder", that identifies  
    EAS as the geocoder that successfully found address Location
  - Merge to final output
* Records that fail EAS geocoder:
  - Are converted to a table, geocoded again using the Street Center Lines 
    geocoder. Records that pass get the same add field treatment as described above
    and are labeled "SC" in the geocoder field
* Records that fail Street Centerlines geocoder:
   - Are added to the feature class generated during the Street Centerlines 
     process, and are labeled "U" in geocoder field
* Merge of feature class(es), generating a final feature class named "final"

*******************************************************************************

DEPENDENCIES:

  * arcpy
  * openpyxl

*******************************************************************************

FILE STRUCTURE:

1 -> imperfect.addresses.py

  * Script that parses through excel file and transform addressees
  * Dependencies are openpyxl, numerate_xl.py

2 -> numerate_xl.py

  * Stores dictionary that reads excel columns as numbers
  * Function that looks for target columns

3 -> geocoding.py

  * Geoprocessing script generated with ArcGIS model builder
  * Dependency is arcpy

4 -> automated_geocoder.py

  * Trigger file that batch processes all scripts
  * Dependencies are all files in folder structure and arcpy

*******************************************************************************

CONSTRAINS AND POTENTIAL ISSUES

* Syntax of all scripts observes Python3 rules, EXCEPT for 
  automated_geocoder.py, which calls for user input with strict Python2
  syntax. If environment is uses Python 3, line 9 must be altered to reflect 
  new syntax (raw_input ---> input) 

* Program overwrites all data everytime is processed

* openpyxl 


*******************************************************************************

PRODUCTION AND FEATURES TO BE IMPLEMENTED

* Program crashes when executed directly from Windows, although it finishes processes successfully
* Add try and except to geocoding.py
* REGEX improvements (as edge cases are identified)
* UI improvements?
* Reduce runtime?


*******************************************************************************


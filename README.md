# Automated Geocoder

This program was first developed for the Urban Planning Department of the City and County of San Francisco, to facilitate the mapping of data provided via spreadsheets.

The program first standardizes addresses bases on standards set by the city for the city. It generates a new spreadsheet as output. It then proceeds to geocode(optionally) the newly generated spreadsheet, using two in-house geocoders.

This script can be easily picked apart and modified to serve the same purpose in different environments, or using different datasets/software.


## Last Modified

The final version of this program was released 02/23/2018.
Subsequent updates may follow as needed.

## Functionality

1 -> Wipes out and compacts target geodatabase (I:\GIS\OASIS\Geocoder\geocoder.gdb)


2 -> Parses through Excel spreadsheet looking for address data

	* Column must be named "Address", "address", "Location" or "location"
	* Target column name above must be unique
	* Excel file must be named "test.xlsx"
	* .xls files are not accepted

3 -> Standardizes address input observing the Enterprise Address System database format

 * Address types are abbreviated, following the standards found here:

 https://data.sfgov.org/Geographic-Locations-and-Boundaries/Street-Names/6d9h-4u5v/data

 * Everything to the right of an address type is wiped out (as in
     "123 1st street, Apt 345"  --> "123 1st st"
 * Apostrophes are removed and leave no whitespace in between characters
 * Ranges (as in "517-520 1st Street") are eliminated. Only the first
   number in the range is kept
 * Removes fractions in addresses (as in 10/2 Market Street)
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
 * Transforms BAYSHORE BLVD --> BAY SHORE BLVD
 * Handles Embarcadero Center Transformations as:
      1 embarcadero center --> 301 CLAY ST
      2 embarcadero center --> 201 CLAY ST
      3 embarcadero center --> 101 CLAY ST
      4 embarcadero center --> 150 DRUMM ST
 * Handles addresses commonly entered without address type
 * Handles assorted edge cases as they are identified

4 -> Generates new spreadsheet

* Writes new column to the next available column in original spreadsheet
* Saves workbook changes and writes address transformations
* New spreadsheet generated is named "transformed.xlsx"

5 -> Geocodes spreadsheet (optional)

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
     geocoding process, and are labeled "U" in geocoder field
* Merge of feature class(es), generating a final feature class named "final"


## DEPENDENCIES:

  * arcpy
  * openpyxl


## FILE STRUCTURE:

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
  * Dependencies are all files in folder structure, openpyxl and arcpy


## CONSTRAINS AND POTENTIAL ISSUES

* Syntax of all scripts observes Python3 rules, EXCEPT for
  automated_geocoder.py, which calls for user input with strict Python2
  syntax. If environment uses Python3, line 50 must be altered to reflect
  new syntax (raw_input ---> input)

* Program overwrites all data everytime is processed

* openpyxl



## PRODUCTION AND FEATURES TO BE IMPLEMENTED

* Program crashes when executed directly from Windows, although it finishes
  processes successfully

IMPLEMENTED --> Add try and except to automated_geocoder

IMPLEMENTED --> UI improvements?

IMPLEMENTED --> It's easy to lock the geodatabase (and locking will be more
                likely once multiple people are using it) - so could we add
                a check for the lock and then print a message saying that it's
                locked and then end the script, rather than it just crashing
                if it's locked?


IMPLEMENTED -->  Maybe print a message at the end with summary of the results -
                 saying something like:
                 Ended successfully!
     		 Cleaned addresses are stored in the I:\GIS\OASIS\Geocoder\transformed.xlsx spreadsheet
	         Results are in this geodatabase: I:\GIS\OASIS\Geocoder\geocoder.gdb
	         Geocoded feature class is called Final. Those not geocoded are flagged with a U in the geocoder field.
 	         X were geocoded by EAS, Y by street centerlines, Z were not geocoded.   95.5% success rate.


IMPLEMENTED -->  It expects the latitude, longitude and zip fields to be numbers (reasonably), but occasionally they are not -
   		 e.g. if there is text in the zip field then it crashes when writing to the failed_table.  Also, sometimes ' '
	         (empty text string) appears to be getting written to the failed_table for the latitude or longitude, which
	         causes it to crash.  Changing the zip, latitude and longitude fields to text in the failed_table and the final merge
	         might fix this...or possibly do some check (select followed by field calc) to change the empty strings to nulls and
	         removing any non-numeric characters from the fields.  
 	         "I:\GIS\OASIS\Requests\2018\20180104_Geocode\Businesses_Clean.xlsx" is a good one for testing this.

## Authors

* **Jonathan Engelbert** - *Sole Developer* - [jonathanengelbert](https://github.com/jonathanengelbert/)

This project has no participating contributors at the moment.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* To all the amazing resources available online
* Mike Wynne
* Python Documentation (https://docs.python.org/3/)

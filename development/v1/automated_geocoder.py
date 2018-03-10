#==============================================================================

#AUTOMATED GEOCODER

# Last Modified: 02/23/2018
# Author: Jonathan Engelbert (Jonathan.Engelbert@sfgov.org)

# Description: This script calls scripts that standardizes and geocodes a list
#  of addresses from an Excel spreadsheet


#==============================================================================

import openpyxl
import imperfect_addresses
import geocoding

print("\t\tLoading Spreadsheet...")

standardize_addresses = True
print("\t\t\nSPREADSHEET LOADED")
geocode = raw_input("\n\tWould you like to geocode the "
                             "spreadsheet?\n(Y/N)\n")

if geocode == "Y" or geocode == "y":
    geocode_spreadsheet = True

else:
    geocode_spreadsheet = False
    print("\nExiting...")

if standardize_addresses:
    imperfect_addresses.transform()

if geocode_spreadsheet:
    print("\nGeocoding addresses...")
    geocoding.geocode()
    print("\nGEOCODING SUCCESSFUL")

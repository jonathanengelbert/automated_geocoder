#==============================================================================
#AUTOMATED GEOCODER

# Last Modified: 02/27/2018
# Author: Jonathan Engelbert (Jonathan.Engelbert@sfgov.org)

# Description: This script calls scripts that standardizes and geocodes a list
#  of addresses from an Excel spreadsheet


#==============================================================================

import openpyxl
import imperfect_addresses
import geocoding_v2
import os

print("======================================================")
print("=================AUTOMATED GEOCODER===================")
print("======================================================")

#PATHS

gdb = "I:\GIS\OASIS\Geocoder\geocoder.gdb"
arcpy.env.workspace = gdb

#DELETE EXISTING FILES IN GEODATABASE

print("\nPreparing Geodatabase...")

feature_classes = arcpy.ListFeatureClasses()
tables = arcpy.ListTables()

try:
    for fc in feature_classes:
        arcpy.Delete_management(fc)

    for table in tables:
        arcpy.Delete_management(table)

#COMPRESS GEODATABASE

    arcpy.CompressFileGeodatabaseData_management(gdb)

    print("GEODATABASE READY")

except Exception as e:
    print("FAILED TO PREPARE GEODATABASE. POSSIBLE LOCK. \n\n")
    print(e)
    raw_input()

#INITIALIZING SPREADSHEET AND SWITCHES

standardize_addresses = True

geocode = raw_input("\n\tWould you like to geocode the "
                             "spreadsheet?\n\t\t\t(Y/N)\n")

if geocode == "Y" or geocode == "y":
    geocode_spreadsheet = True

else:
    geocode_spreadsheet = False

#PROCESSES

if standardize_addresses:
    try:
        imperfect_addresses.transform()
        print("\n******************************************************\n")
        print("Cleaned addresses are stored in:\nI:\GIS\OASIS\Geocoder"
              "\\transformed.xlsx")
        print("\n******************************************************\n")
        if geocode_spreadsheet:
            try:
                print("Geocoding addresses...")
                geocoding_v2.geocode()

#GENERATES REPORT FOR GEOCODING

                # Cursor and target feature class

                fc = "I:\\GIS\\OASIS\\Geocoder\\geocoder.gdb\\final"
                cursor = arcpy.da.SearchCursor(fc, ['geocoder'])

                # Variables

                eas = 0
                sc = 0
                u = 0
                total = arcpy.GetCount_management(fc).getOutput(0)

                # Logic

                for row in cursor:
                    if "EAS" in row:
                        eas += 1
                    elif "SC" in row:
                        sc += 1
                    elif "U" in row:
                        u += 1

                # Result Rates

                eas_percentage = (100 * eas / int(total))
                sc_percentage = (100 * sc / int(total))
                u_percentage = (100 * u / int(total))
                success_rate = eas_percentage + sc_percentage

                #Report Output:

                print(
                    "\n******************************************************\n")
                print("GEOCODING RESULTS:\n\nMaster Address Geocoder: " +
                      str(eas) + " record(s) geocoded(" + str(eas_percentage)
                      + "%)")

                print("Street Centerlines Geocoder: " + str(sc) + " "
                                                                     "record(s) "
                    "geocoded(" + str(sc_percentage) + "%)")

                print("Unmatched Records: " + str(u) + " "
                "record(s) "
                "not geocoded(" + str(u_percentage) + "%)")

                print("\n\nRECORDS GEOCODED: " +
                      str(success_rate) + "%")

                print(
                    "\n******************************************************\n")

                raw_input()
            except Exception as e:
                print("\nGEOCODING FAILED\n\n")
                print(e)

    except Exception as e:
        print("\nADDRESS TRANSFORMATION FAILED\n\n")
        print(e)


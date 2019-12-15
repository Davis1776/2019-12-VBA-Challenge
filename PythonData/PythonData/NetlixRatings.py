# First we'll import the os module
# This will allow us to create file paths across operating systems
import os

# Module for reading CSV files
import csv

NetFlix_File = 'netflixratings.csv'

# Improved Reading using CSV module

with open(NetFlix_File, 'r', newline='') as csvfile: netflixratings.csv

    # CSV reader specifies delimiter and variable that holds contents
    csvreader = csv.reader(csvfile, delimiter=',')
    print(csvreader)
    
    # Read the header row first (skip this step if there is now header)
    csv_header = next(csvreader)
    print(f"CSV Header: {csv_header}")
    
    # Read each row of data after the header
    for row in csvreader:
        print(row)
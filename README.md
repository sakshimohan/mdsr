# Cleaning Maternal Death Surveillance Report Data (MDSR) for Malawi

MDSR Excel files are currently in excel formats which are not easy to analyse. This is an attempt at writing a generalisable code to parse
through the excel files to generate a clean dataset. Certain inputs need to be provided for the function to work properly - row in which data begins, column in which data begins, number of districts. 

_parse_mdsr_files.py_ defines the function `cleanmdsrdta` to extract relevant data from MDSR files submitted by districts. 
_generate_mdsr_database.py_ demonstrates how the 2018 national data was put together using `cleanmdsrdta`for 2018. 

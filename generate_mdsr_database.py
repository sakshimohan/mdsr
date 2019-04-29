# Set file path

path = '/Users/sakshimohan/Dropbox (Personal)/MOH Malawi/Data/MDSR (Maternal Death Surveillance Report)/processing'
os.chdir(path)
os.getcwd()

#  GENERATE DATAFRAMES FOR RELEVANT ZONES AND TIME PERIODS

## CENTRAL EAST ZONE

cez_janmar = cleanmdsrdta('Jan-Mar 2018 - CEZ.xlsx', "MDSR DB - All Zones", 9, 4, 5, 'cez_janmar')
cez_aprjun = cleanmdsrdta('Apr-Jun 2018 - CEZ.xlsx', "MDSR DB - All Zones", 9, 4, 5, 'cez_aprjun')
cez_julsep = cleanmdsrdta('Jul-Sep 2018 - CEZ.xlsx', "MDSR DB - All Zones", 9, 4, 5, 'cez_julsep') 
cez_octdec = cleanmdsrdta('Oct-Dec 2018 - CEZ.xlsx', "MDSR DB - All Zones", 9, 56 ,5 , 'cez_octdec')

## CENTRAL WEST ZONE

cwz_janmar = cleanmdsrdta('Jan-Mar 2018 - CWZ.xlsx', "MDSR DB - All Zones", 9, 36, 4 , 'cwz_janmar')
cwz_aprjun = cleanmdsrdta('Apr-Jun 2018 - CWZ.xlsx', "MDSR DB - All Zones", 9, 36, 4 , 'cwz_aprjun')
cwz_julsep = cleanmdsrdta('Jul-Sep 2018 - CWZ.xlsx', "MDSR DB - All Zones", 9, 36, 4 , 'cwz_julsep') 
cwz_octdec = cleanmdsrdta('Oct-Dec 2018 - CWZ.xlsx', "MDSR DB - All Zones", 9, 36, 4 , 'cwz_octdec')

## NORTH ZONE

nz_janmar = cleanmdsrdta('Jan-Mar 2018 - NZ.xlsx', "MDSR DB - North", 8, 4, 7, 'nz_janmar')
nz_aprjun = cleanmdsrdta('Apr-Jun 2018 - NZ.xlsx', "MDSR DB - All Zones", 8, 4, 7, 'nz_aprjun')
nz_julsep = cleanmdsrdta('Jul-Sep 2018 - NZ.xlsx', "MDSR DB - All Zones", 8, 4, 7, 'nz_julsep') 
nz_octdec = cleanmdsrdta('Oct-Dec 2018 - NZ.xlsx', "MDSR DB - North", 8, 4, 7, 'nz_octdec')

## SOUTH EAST ZONE

#sez_janmar = cleanmdsrdta('Jan-Mar 2018 - SEZ.xlsx', "MDSR DB - All Zones", 8, 4, 7, 'sez_janmar')
sez_aprjun = cleanmdsrdta('Apr-Jun 2018 - SEZ.xlsx', "MDSR DB - All Zones", 8, 4, 6, 'sez_aprjun')
sez_julsep = cleanmdsrdta('Jul-Sep 2018 - SEZ.xlsx', "MDSR DB - All Zones", 8, 4, 6, 'sez_julsep') 
sez_octdec = cleanmdsrdta('Oct-Dec 2018 - SEZ.xlsx', "Oct-Dec118", 8, 4, 7, 'sez_octdec')

## SOUTH WEST ZONE

swz_janmar = cleanmdsrdta('Jan-Mar 2018 - SWZ.xlsx', "MDSR DB - All Zones", 8, 4, 7, 'swz_janmar')
swz_aprjun = cleanmdsrdta('Apr-Jun 2018 - SWZ.xlsx', "MDSR DB - All Zones", 8, 4, 7, 'swz_aprjun')
swz_julsep = cleanmdsrdta('Jul-Sep 2018 - SWZ.xlsx', "MDSR DB - South West", 8, 4, 7, 'swz_julsep') 
swz_octdec = cleanmdsrdta('Oct-Dec 2018 - SWZ.xlsx', "MDSR DB - All Zones", 8, 4, 7, 'swz_octdec')

columns = ["Zone", "District"]

# Generate dataframe to match districts to zones
zone_mapping = [['Central East','Nkhotakota'],
                ['Central East','Ntchisi'],
                ['Central East','Salima'],
                ['Central East','Dowa'],
                ['Central East','Kasungu'],
                ['Central West','Dedza'],
                ['Central West','Lilongwe'],
                ['Central West','Mchinji'],
                ['Central West','Ntcheu'],
                ['North','Chitipa'],
                ['North','Karonga'],
                ['North','Likoma'],
                ['North','Mzimba North'],
                ['North','Mzimba South'],
                ['North','Nkhata Bay'],
                ['North','Rumphi'],
                ['South East','Machinga'],
                ['South East','Mangochi'],
                ['South East','Mulanje'],
                ['South East','Zomba'],
                ['South East','Balaka'],
                ['South East','Zomba Central Hospital'],
                ['South East','Phalombe'],
                ['South West','Chiradzulu'],
                ['South West','Nsanje'],
                ['South West','Blantyre'],
                ['South West','Chikwawa'],
                ['South West','Mwanza'],
                ['South West','Neno'],
                ['South West','Thyolo'],
                ['South West','Blantyre (No QECH data)']]
zone_mapping = pd.DataFrame(zone_mapping, columns = columns)

# APPEND ALL DATAFRAMES AND MERGE ZONAL MAPPING

full_data = cez_janmar
full_data = full_data.append(cez_aprjun, ignore_index = False)
full_data = full_data.append(cez_julsep, ignore_index = False)
full_data = full_data.append(cez_octdec, ignore_index = False)
full_data = full_data.append(cwz_janmar, ignore_index = False)
full_data = full_data.append(cwz_aprjun, ignore_index = False)
full_data = full_data.append(cwz_julsep, ignore_index = False)
full_data = full_data.append(cwz_octdec, ignore_index = False)
full_data = full_data.append(nz_janmar, ignore_index = False)
full_data = full_data.append(nz_aprjun, ignore_index = False)
full_data = full_data.append(nz_julsep, ignore_index = False)
full_data = full_data.append(nz_octdec, ignore_index = False)
#full_data = full_data.append(sez_janmar, ignore_index = False)
full_data = full_data.append(sez_aprjun, ignore_index = False)
full_data = full_data.append(sez_julsep, ignore_index = False)
full_data = full_data.append(sez_octdec, ignore_index = False)
full_data = full_data.append(swz_janmar, ignore_index = False)
full_data = full_data.append(swz_aprjun, ignore_index = False)
full_data = full_data.append(swz_julsep, ignore_index = False)
full_data = full_data.append(swz_octdec, ignore_index = False)

full_data_merge = zone_mapping.merge(full_data, how='outer', on = 'District')
full_data_merge.to_csv("full_data.csv", header=True)

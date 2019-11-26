"""
    This file converts the coordinates from a KML file obtained from Google Earth to an excel file
    Coordinates are saved in the format of DDMMSS.SS

    Copyright (C) 2019 Bakthakolahalan Shyamsundar

    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""

"""
Instructions:
-------------
* This program depends on pandas and Beautiful soup libraries. Make sure to install them.

* Change the filename and filepath to the name of the KML file (without the extension) and its directory

* Change the type inorder to get appropriate suffix in the filenames

"""

from bs4 import BeautifulSoup
import pandas as pd

def main():
    filename = ''
    filepath = ''

    """
    Select type as follows:
        't' for Takeoff Zone
        'r' for Recovery Zone
        'g' for Geofence
    """
    type = 'r'

    if type == 'g':
        suffix = '_Airspace.kml'
        xlsuffix = '_geofence_coordinates.xlsx'
        print('Converting geofence coordinates to Excel......')
    elif type == 't':
        suffix = '_Takeoff_Zone.kml'
        xlsuffix = '_takeoff_zone_coordinates.xlsx'
        print('Converting takeoff zone coordinates to Excel......')
    elif type == 'r':
        suffix = '_Recovery_Zone.kml'
        xlsuffix ='_recovery_zone_coordinates.xlsx'
        print('Converting recovery zone coordinates to Excel......')
    else:
        print('Invalid type')

    with open(filepath+filename+suffix, 'r') as f:
        s = BeautifulSoup(f, 'xml')
        for coords in s.find_all('coordinates'):
            latlongalt = coords.string.strip().split(" ")
            df = pd.DataFrame(columns=['Latitude','Longitude']);
            print('File location and name:' + filepath + filename + xlsuffix)
            writer = pd.ExcelWriter(filepath + filename + xlsuffix)

            i = 0
            for split in latlongalt[0:]:
                temp = []
                latlongtemp = split.split(',')

                lat = float(latlongtemp[1])
                long = float(latlongtemp[0])

                lat_out_deg = int(lat)
                lat_out_min = int((lat - lat_out_deg) * 60)
                if(lat_out_min < 10):
                    lat_out_min = '0' + str(lat_out_min)
                lat_out_sec = (- int(lat_out_min) + (lat - int(lat))*60)* 60
                if(lat_out_sec < 10):
                    lat_out_sec = '0' + str(lat_out_sec)
                temp.append(str(lat_out_deg) + str(lat_out_min) + str(lat_out_sec)[0:5])

                long_out_deg = int(long)
                long_out_min = int((long - long_out_deg) * 60)
                if(long_out_min < 10):
                    long_out_min = '0' + str(long_out_min)
                long_out_sec = (- int(long_out_min) + (long - int(long))*60)* 60
                if(long_out_sec < 10):
                    long_out_sec = '0' + str(long_out_sec)
                temp.append(str(long_out_deg) + str(long_out_min) + str(long_out_sec)[0:5])

                df.loc[i] = temp
                i = i+1

            df.to_excel(writer,'Sheet1',index=False)
            writer.save()

            print('KML to Excel conversion successful. xlsx file created. ')


if __name__ == "__main__":
    main()

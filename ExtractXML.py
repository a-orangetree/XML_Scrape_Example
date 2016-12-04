import pandas as pd
import sqlite3
from sqlalchemy import create_engine
import os
from lxml import objectify


# Enter paths to files here
pathToXML = 'c:\\automation\LayoutsData\\'
pathToSQLiteDB = 'c:\\automation\Automation.db'


#Create connection to SQLite database
conn = sqlite3.connect(pathToSQLiteDB)
cur = conn.cursor()


#Drop old tables in SQLite database, if they exist
cur.executescript('''

DROP TABLE IF EXISTS XML;
DROP TABLE IF EXISTS XML2;
                      
''')

print ('Tables dropped, if needed')



#Connect to database engine to load data in SQLite
engine = create_engine('sqlite:///'+pathToSQLiteDB)

#Connect to XML data
pathToXMLFiles = pathToXML
XMLDirectory = os.listdir(pathToXMLFiles)


#Set AutoID for use below
autoID = int(0)


#Creates empty Pandas Dataframes for use below
xmlDF = pd.DataFrame(columns = ['AutoID','FileName','Id','Position','PositionState','ProcessId','Type',\
                                'WindowTitle','ZOrder','AudioStatus','ConferenceCallStatus',\
                                'SourceObject','VisibleChrome'])

xmlDF2 = pd.DataFrame(columns = ['AutoID','FileName','CanvasID','Height','Width','X','Y','IsAudioOn',\
                                 'IsMicrophoneOn','CallerStatus','ObjectID','ObjectName','ObjectType'])


#Loops through each XML file to extract data
for file in XMLDirectory:
    xmlPath = pathToXML+file
    xml = objectify.parse(open(xmlPath))
    root = xml.getroot()
   
    
#Extract data from XML and save in a Pandas Dataframe    
    for i in range(len(root.getchildren())):
        obj = root.getchildren()[i].getchildren()
        row = dict(zip(['AutoID','FileName','Id','Position','PositionState','ProcessId','Type','WindowTitle',\
        'ZOrder','AudioStatus','ConferenceCallStatus','SourceObject','VisibleChrome']\
        ,[str(autoID), str(file)[:-4],obj[0].text, obj[1].text,obj[2].text,obj[3].text,obj[4].text,obj[5].text,\
        obj[6].text,obj[7].text,obj[8].text,obj[9].text,obj[10].text]))
        row_s = pd.Series(row)
        row_s.name = i
        xmlDF = xmlDF.append(row_s)
        
        canvasID = root.getchildren()[i].getchildren()[0]
        height = root.getchildren()[i].getchildren()[1].getchildren()[0]
        width = root.getchildren()[i].getchildren()[1].getchildren()[1]
        xTile = root.getchildren()[i].getchildren()[1].getchildren()[2]
        yTile = root.getchildren()[i].getchildren()[1].getchildren()[3]
        isAudioOn = root.getchildren()[i].getchildren()[8].getchildren()[0]
        isMicrophoneOn = root.getchildren()[i].getchildren()[8].getchildren()[1]
        callerStatus = root.getchildren()[i].getchildren()[8].getchildren()[2]
        objectID = root.getchildren()[i].getchildren()[9].getchildren()[0]
        objectName = root.getchildren()[i].getchildren()[9].getchildren()[1]
        objectType = root.getchildren()[i].getchildren()[9].getchildren()[2]

        row2 = dict(zip(['AutoID','FileName','CanvasID','Height','Width','X','Y','IsAudioOn','IsMicrophoneOn',\
        'CallerStatus','ObjectID','ObjectName','ObjectType'],\
        [str(autoID), str(file)[:-4], canvasID.text, height.text, width.text, xTile.text, yTile.text,\
         isAudioOn.text, isMicrophoneOn.text, callerStatus.text, objectID.text,\
         objectName.text, objectType.text]))
        row_s2 = pd.Series(row2)
        row_s2.name = i
        xmlDF2 = xmlDF2.append(row_s2)
        
    autoID = autoID+1

    
#Push Dataframes containing XML data to SQLite database
xmlDF.to_sql('XML',engine)
xmlDF2.to_sql('XML2',engine)

print ('Table created from XML')

    
print('Script finished')

#Close connection to database
conn.close()
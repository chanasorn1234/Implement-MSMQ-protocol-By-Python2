import re
x = '<Root xmlns:dt="urn:schemas-microsoft-com:datatypes">' +\
	    '<Dictionary key="Top">'+\
	        '<Dictionary key="BinIndex1">'+\
	            '<V dt:dt="string" key="BinNumber">0</V>'+\
	            '<V dt:dt="string" key="BinDescription">System Error</V>'+\
	            '<V dt:dt="string" key="BinGrade">F</V>'+\
	        '</Dictionary>'+\
            '<V dt:dt="string" key="BinCount">1</V>'+\
	        '<V dt:dt="string" key="LoadboardID">N/A</V>'+\
	        '<V dt:dt="string" key="ContactorID">N/A</V>'+\
	        '<V dt:dt="i4" key="ReturnCode">0</V>'+\
	        '<V dt:dt="string" key="ReturnText"></V>'+\
	        '<V dt:dt="string" key="ReturnDetails"></V>'+\
	        '<V dt:dt="string" key="SetupTesterMsgID">{1747890D-80E9-4A4B-B1D2-3901143B0B68}\8555861</V>'+\
	        '<V dt:dt="string" key="SenderID">T009SPEA</V>'+\
	    '</Dictionary>'+\
	'</Root>'

y = x.replace(" ","")
y = re.sub('"string"','"string" ',y)
y = re.sub("<Root","<Root ",y)
y = re.sub("<Dictionary","<Dictionary ",y)
y = re.sub("<V","<V ",y)
y = re.sub('"i2"','"i2" ',y)
y = re.sub('"i4"','"i4" ',y)
print(x)
print()
print(y)



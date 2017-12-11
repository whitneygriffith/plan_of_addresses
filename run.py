'''
Problem: 

One item in each sheet gets left out. For instance ANC 1 has 28 individuals but only 27 gets converted.

Querying Google API too many times gets a timeout.any
'''

import xlrd
import xlwt
import googlemaps
from datetime import datetime
from config import gmaps
import time

workbook_name = str(raw_input('What is the name of the excel? (Cases matter): '))
workbook = xlrd.open_workbook(workbook_name)

#Use input() if python 3
sheet = str(raw_input('Enter sheet name (Cases matter): '))

#Function to get specific sheet
def get_sheet(sheet):
    try:
        sheet = workbook.sheet_by_name(sheet)
        return sheet
    except:
        print("Not a valid sheet name")
        return

anc = get_sheet(sheet)

#Loop until valid sheet is found
while (anc == None):
    sheet = str(raw_input('Enter sheet name: '))
    anc = get_sheet(sheet)




#store values in a list nested with a dictionary list = [{last: '', first: '', lat:'', long:'', },]

#Desired Columns: 3,4,6,7
#Desired Rows: 1 until sheet is empty


#Counter for row .cell(row, col)
i = 1
#placeholders for desired column
col1 = 3
col2 = 4
col3 = 6
col4 = 7

#create unsorted list
unsorted_plan = []

#loop until n rows; adds all the info into a list
for i in range(1, anc.nrows):
    dict1 = {'Permit #': anc.cell(i, 0).value, 'Issued': anc.cell(i, 1).value, 'Expiration': anc.cell(i, 2).value, 'last': anc.cell(i, 3).value, 'first': anc.cell(i, 4).value, 'Address': anc.cell(i, 5).value, 'lat': anc.cell(i, 6).value, 'longitude': anc.cell(i, 7).value, 'ANC': anc.cell(i, 8).value, 'Ward': anc.cell(i, 9).value}
    unsorted_plan.append(dict1)
    i += 1

#create sorted list
plan = []
plan.append({'Permit #': '1-33', 'Issued': '', 'Expiration': '40117', 'last': 'Purcell', 'first': 'George', 'Address': '1818 8TH STREET NW', 'lat': '38.91460307', 'longitude': '-77.02319348', 'ANC': 'ANC 1B', 'Ward': 'Ward 1', 'transit_time': '1433'}) 

									
#Set starting point as DDOT (38.876256, -77.00699) 
start = (38.876256, -77.00699)
#Value to store time of transit
transit_time = 0

#Ask desired mode of transport
mode1 = str(raw_input('Enter desired mode of transport. Options are bicycling, walking, transit, driving (Cases matter): '))


#origlength of list
length = len(unsorted_plan)


#point variable
other_points = {}

count = 2

#Create results spreadsheet
resultbook = xlwt.Workbook()
sheet = resultbook.add_sheet('Sheet1')
#loop until len(plan) == len(unsortedPlan)
while len(plan) != length:
    #use the last item in plan as the start
    index = int(len(plan) - 1)
    lat1 = plan[index]['lat']
    longitude1 = plan[index]['longitude']
    start1 = (lat1, longitude1)
    #Get point closest to start
    for i in range(0, len(unsorted_plan)):
        lat = unsorted_plan[i]['lat']
        longitude = unsorted_plan[i]['longitude']
        end = (lat, longitude)
        time.sleep(2) 
        try: 

            directions_result = gmaps.directions(start1,
                                                end,
                                                mode=mode1,
                                                departure_time=datetime.now())
        except:
            time.sleep(90)  # sleep for 90 seconds
            directions_result = gmaps.directions(start1,
                                    end,
                                    mode=mode1,
                                    departure_time=datetime.now())


        #temp duration
        # value gets you transition value in seconds
        temp = directions_result[0]['legs'][0]['duration']['value']
        #reset the transit time for the i in dictionary
        unsorted_plan[i]['transit_time'] = temp
        if i == 0:
            transit_time = temp
            other_points = unsorted_plan[i]
            
        if temp <= transit_time:
            transit_time = temp
            other_points = unsorted_plan[i]
            
    #Add other_points to sorted list
    plan.append(other_points)

    sheet.row(count).write(0, other_points["Permit #"])
    sheet.row(count).write(1, other_points["Issued"])
    sheet.row(count).write(2, other_points["Expiration"])
    sheet.row(count).write(3, other_points["last"])
    sheet.row(count).write(4, other_points["first"])
    sheet.row(count).write(5, other_points["Address"])
    sheet.row(count).write(6, other_points["lat"])
    sheet.row(count).write(7, other_points["longitude"])
    sheet.row(count).write(8, other_points["ANC"])
    sheet.row(count).write(9, other_points["Ward"])
    sheet.row(count).write(10,other_points["transit_time"])
    count += 1
    resultbook.save("results2.xls")
    print ("completed ", count)
    #Delete injected person from unsorted list
    if other_points in unsorted_plan:
        unsorted_plan.remove(other_points)






'''
time.pause(86400)
#origlength of list
length = len(unsorted_plan)



#reset transit_time
transit_time = 0
#point variable
other_points = {}

#loop until len(plan) == len(unsortedPlan)
while len(plan) != length:
    #use the last item in plan as the start
    index = int(len(plan) - 1)
    lat1 = plan[index]['lat']
    longitude1 = plan[index]['longitude']
    start1 = (lat1, longitude1)
    #Get point closest to start
    for i in range(0, len(unsorted_plan)):
        lat = unsorted_plan[i]['lat']
        longitude = unsorted_plan[i]['longitude']
        end = (lat, longitude)
        time.sleep(2) 
        try: 

            directions_result = gmaps.directions(start1,
                                                end,
                                                mode=mode1,
                                                departure_time=datetime.now())
        except:
            time.sleep(90)  # sleep for 90 seconds
            directions_result = gmaps.directions(start1,
                                    end,
                                    mode=mode1,
                                    departure_time=datetime.now())


        #temp duration
        # value gets you transition value in seconds
        temp = directions_result[0]['legs'][0]['duration']['value']
        #reset the transit time for the i in dictionary
        unsorted_plan[i]['transit_time'] = temp
        if i == 0:
            transit_time = temp
            other_points = unsorted_plan[i]
            
        if temp <= transit_time:
            transit_time = temp
            other_points = unsorted_plan[i]
            
    #Add other_points to sorted list
    plan.append(other_points)
    #Delete injected person from unsorted list
    if other_points in unsorted_plan:
        unsorted_plan.remove(other_points)





#Create results spreadsheet
resultbook = xlwt.Workbook()
sheet = resultbook.add_sheet('Sheet1')
cols = ["Permit  #", "Issued",	"Expiration",	"Last",	"First",	"Address",	"Lattitude",	"Longitude",	"ANC",	"Ward", "Transit Time  (s)", "Transit Time (hours and mins)"]
length = len(cols) 

#Make header
for i in range(length):
    sheet.row(0).write(i, cols[i])

#count to increase row
count = 1


for i in range(0, len(plan)):
    sheet.row(count).write(0, plan[i]["Permit #"])
    sheet.row(count).write(1, plan[i]["Issued"])
    sheet.row(count).write(2, plan[i]["Expiration"])
    sheet.row(count).write(3, plan[i]["last"])
    sheet.row(count).write(4, plan[i]["first"])
    sheet.row(count).write(5, plan[i]["Address])
    sheet.row(count).write(6, plan[i]["lat"])
    sheet.row(count).write(7, plan[i]["longitude"])
    sheet.row(count).write(8, plan[i]["ANC"])
    sheet.row(count).write(9, plan[i]["Ward"])
    sheet.row(count).write(10,plan[i]["transit_time"])
    
  #  sheet.row(count).write(10, sec)
    count +=1


resultbook.save("results.xls")

'''
'''
#print plan to test it
for i in range(0, len(plan)):
    print plan[i]['last']


#Sort the rest of the list: Valid for fixed start location
for i in range(0, len(unsorted_plan)):
    #ignore the item that is point1
    if unsorted_plan[i] == point1:
        continue

    lat = unsorted_plan[i]['lat']
    longitude = unsorted_plan[i]['longitude']
    end = (lat, longitude)
    directions_result = gmaps.directions(start,
                                         end,
                                         mode="walking",
                                         departure_time=now)
    #temp duration
    # value gets you transition value in seconds
    temp = directions_result[0]['legs'][0]['duration']['value']
    #append the transit time to the dictionary
    unsorted_plan[i]['transit_time'] = temp
    #Find place in plan to add this one too

    for u in range(0, len(plan)):
        if temp <= plan[u]['transit_time']:
            if temp >= plan[u - 1]['transit_time']:
                plan.insert(u, unsorted_plan[i])
            break
        elif u == len(plan) -1:
            if temp >= plan[u]['transit_time']:
                plan.append(unsorted_plan[i])

'''
                              





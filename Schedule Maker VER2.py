import PySimpleGUI as sg
import datetime
import requests, bs4, docx

#make document name from current date
dt = datetime.datetime.now()
currentDate = (str(dt.month)+"."+str(dt.day)+"."+str(dt.year))
documentTitle = "Weekly Bulletin "+ currentDate + ".docx"

newDocument = docx.Document()

#note: the sg.Spin() dropdown returns a string, so no str() is needed later on

layout = [[sg.Spin([i for i in range(0,10)], initial_value=0, key="seriesEvents"), sg.Text('Number of Events to skip')],
    [sg.Spin([i for i in range(1,20)], initial_value=1, key="eventNumber"), sg.Text('Number of Events to Grab')],
    [sg.Text("Input website.")],      
    [sg.InputText(key="siteName")],    
    [sg.Submit(), sg.Cancel()]]

window = sg.Window('Weekly Bulletin Maker', layout)    

event, values = window.read()    
window.close()

#get the website, access the <h3> elements you need
res = requests.get(values["siteName"])
res.raise_for_status()
websiteObject = bs4.BeautifulSoup(res.text, features="lxml")
eventLinks = websiteObject.findAll("h3")
links = []
for headers in eventLinks:
    for anchors in headers.findAll("a"):
        links.append(anchors['href'])

#get number of days you iterate over, allowing the option to skip links at the top of the page which usually are excluding from the final document
totalEvents = (values["eventNumber"])+(values["seriesEvents"])

#make data list for outside the for loop


#request each website from the new links list, make objects, get text
for eventSites in links[(values["seriesEvents"]):totalEvents]:
    secondRequest = requests.get(eventSites)
    secondRequest.raise_for_status()
    soup = bs4.BeautifulSoup(secondRequest.text, features='lxml')

    #grabbing element objects
    divElements = soup.find("div", class_="detail")
    dateElements = divElements.p
    locationElements = divElements.a.previous_sibling
    paragraphElements = soup.h3.find_next_sibling("p")
    speakerElement = paragraphElements.find_next_sibling("p")
    

    #turning elements into usable text
    documentHeader = soup.h2.get_text()
    documentSpeaker = speakerElement.get_text(strip=True)
    documentDates = dateElements.get_text(" ", strip=True)
    documentLocation = "Where: "+(locationElements.get_text(strip=True))
    documentParagraphs = paragraphElements.get_text(strip=True)
    
    #document writing
    newDocument.add_heading(documentHeader, 2)
    newDocument.add_paragraph(documentSpeaker)
    newDocument.add_paragraph(documentLocation)
    newDocument.add_paragraph(documentDates)
    newDocument.add_paragraph(documentParagraphs)

newDocument.save(documentTitle)
sg.popup("Success! The document will be found in the same folder where the program is.")


#TODO
###
#!/usr/bin/env python
# coding: utf8

################################################################################
# Programm zum Erstellen von Rechnungen mit Daten aus einer csv und einem Dokument Template.
# Autor: Daniel Strohbach (hallo@daniel-strohbach.de)
# Datum: 05. November 2023
# Infos:
# 1. Das Script wird mit "python generateInvoice.py" ausgeführt - man kann auch dafür eine batch file nutzen
# 2. Es werden eine Datei mit "rechnung-template.svg" und "orders.csv" erwartet.

#Dependencies: 
#pip install lxml

################################################################################

# to do: 
# seitenumbruch bei translate >300 or n>9

#includes
import csv
from ensurepip import version
from itertools import islice
from lxml import etree
from lxml import objectify
import subprocess
import os

#this are the files we need in the same folder
filepath = os.path.dirname(os.path.abspath(__file__))
templateFile = filepath + "/rechnung-template.svg"
inputCSV = filepath + "/orders.csv"

#this is the svg namespace
SVGNS = {"svg":"http://www.w3.org/2000/svg",'inkscape': 'http://www.inkscape.org/namespaces/inkscape'}

# this are the variables inside the template to replace with the data from csv file
platzhalterName = "%Name%"
platzhalterBestellDatum = "%BestellDatum%"
platzhalterBestellNr = "%BestellNr%"
platzhalterRechnungsDatum = "%DatumRechnung%"
platzhalterRechnungsNr = "%NrRechnung%"
platzhalterStrasse = "%Strasse%"
platzhalterPLZ = "%PLZ%"
platzhalterStadt = "%Stadt%"

platzhalterArtikelName = "%Artikel%"
platzhalterArtikelNr = "%ArtikelNr%"
platzhalterAnzahl = "%Anzahl%"
platzhalterPreis = "%Preis%"
platzhalterZwischenSumme ="%ZwischenSumme%"
platzhalterVersandkosten ="%VersandKosten%"
platzhalterGesamtPreis = "%PreisGesamt%"
platzhalterMWST ="%MWST%"

platzhalterSeitenZahl = "%SeitenZahl%"
platzhalterMaxSeitenZahl = "%MaxSeitenZahl%"

#some helper vars
previousCustomer = "none"
n=1 #for rows
p=2 #for multiple pages
translation = 60 #start value
translationX = 800
spacing = 60 # step value
spacingX = 800
ZwischenSumme = 0.0

# as long as the input csv is opened, we do stuff
with open(inputCSV) as csvfile:
    reader = csv.reader(csvfile)    

    #we loop through the csv, skipping the first row with islice (first row is just titles for the columns)   
    for row in islice(reader, 1, None):

        invoiceFile = filepath + "/invoice/invoice-temp-" + row[17] + ".svg"

        #this is just giving us a hint where we are atm
        print("Aktuelle Zeile: ")
        print("Bestelldatum: "+ row[0], "Artikelname: "+ row[1], "Kunde: " + row[17])    

        #if there is a new customer, we start with a fresh document - therefore we reset our counter to 1
        if previousCustomer != row[17]:
            n = 1
            p = 2
            translation = 60 # this is the amount of space the next article Row needs to be offsetet below the previous one. we add them app during the loop to stack the rows
            translationX = 800
            ZwischenSumme = 0.0 #also we need to reset this

        #we go through this path only on first looping
        if n == 1:
            print("n gleich 1")

            #open the invoice template with placeholders
            f = open(templateFile, 'r')
            filedata = f.read()
            f.close()

            # Fill in Adress
            newdata = filedata.replace(platzhalterName,row[17])
            print("replaced Name")
            newdata = newdata.replace(platzhalterStrasse,row[18])
            print("replaced Strasse")
            newdata = newdata.replace(platzhalterPLZ,row[22])
            print("replaced PLZ")
            newdata = newdata.replace(platzhalterStadt,row[20])
            print("replaced Stadt")

            #Fill in Order Dates
            newdata = newdata.replace(platzhalterBestellDatum,row[0])
            print("replaced BestellDatum")
            newdata = newdata.replace(platzhalterBestellNr,row[24])
            print("replaced BestellNr")
            newdata = newdata.replace(platzhalterRechnungsDatum,row[0])
            print("replaced Rechnungsdatum")
            newdata = newdata.replace(platzhalterRechnungsNr,row[24])
            print("replaced RechnungsNr")

            #Fill in Article Stuff that is static
            newdata = newdata.replace(platzhalterVersandkosten,row[9])
            print("replaced Versandkosten zu: "+ str(format(float(row[9]), '.2f'))) 

        #we do this on every loop
        #to do some maths, we have to convert the csv numbers from strings to actuals floats
        Preis = float(row[11])
        print("ArtikelName: "+ str(row[1]))
        print("ArtikelPreis: " + str(Preis))
        ZwischenSumme = ZwischenSumme + Preis #To calculate the Order Sum, we just add them up during looping
        print("ZwischenSumme: " + str(ZwischenSumme))
        Versandkosten = float(row[9])
        print("Versandkosten: " + str(Versandkosten))

        GesamtPreis = ZwischenSumme + Versandkosten # this we just update with every loop, no problem with that :)
        MWST = GesamtPreis * 0.19
        print("GesamtPreis: " + str(GesamtPreis))
        print("MWST: " + str(MWST))  
                
        #this path we take from loop passthrough 2
        if n > 1: 
            print("n größer 1")

            #open last saved invoice to rewrite to it (first read ofc)
            f = open(invoiceFile, 'r')
            filedata = f.read()
            f.close()

            #get element tree from last saved invoice file now
            parser=etree.XMLParser(encoding='UTF-8') #i had a lot of parsing errors because of german umlauts in utf-8
            newtree = etree.parse(invoiceFile, parser)
            root = newtree.getroot()

            # here we loop through the element tree with xpath to find our last added article row to get its "position" in the element tree
            for element in newtree.xpath('//*[local-name()="svg"]//*[local-name()="g"]', namespaces=SVGNS):

                if element.attrib['id'] == "ArticleRow" + str(n-2):
                    #print("ArticleRow for Position detected in existing invoice element tree: " + str(element.attrib['id']))
                    pos = element
                    parent = pos.getparent()
                    #print("found parent with ID: " + parent.attrib['id'])
                    #print()
                
                # #also we will move the section with the sums along with the size of the table - decided not to do it, but position it on bottom of template
                # if element.attrib['id'] == "Summen":
                #     print("found element with ID and move it down: " + element.attrib['id'])
                #     element.attrib['transform'] = "translate(0," + str(translation-spacing) +")"   

            # here we loop through the element tree with xpath to find the attributes of the Sum Stuff to exchange
            for element in newtree.xpath('//*[local-name()="svg"]//*[local-name()="text"]', namespaces=SVGNS):

                if element.attrib ['id'] == "ZwischenSumme":
                    #print("TextElement detected in existing invoice element tree: " + str(element.attrib['id']))

                    for span in element:
                        #print("Children of TextElement detected: " + str(span.attrib['id']))

                        for child in span:
                            #print("Children of Child detected: " + str(child.attrib['id']))
                            child.text = str(format(ZwischenSumme, '.2f'))
                            #print("replaced ZwischenSumme zu: "+ str(format(ZwischenSumme, '.2f')))
                
                if element.attrib ['id'] == "GesamtPreis":
                   # print("TextElement detected in existing invoice element tree: " + str(element.attrib['id']))

                    for span in element:
                        #print("Children of TextElement detected: " + str(span.attrib['id']))

                        for child in span:
                            #print("Children of Child detected: " + str(child.attrib['id']))
                            child.text = str(format(GesamtPreis, '.2f'))
                            #print("replaced GesamtPreis to: "+ str(format(GesamtPreis, '.2f')))
                
                if element.attrib ['id'] == "MWST":
                    #print("TextElement detected in existing invoice element tree: " + str(element.attrib['id']))

                    for span in element:
                        #print("Children of TextElement detected: " + str(span.attrib['id']))

                        for child in span:
                            #print("Children of Child detected: " + str(child.attrib['id']))
                            child.text = str(format(MWST, '.2f'))
                            #print("replaced MWST to: "+ str(format(MWST, '.2f')))
                
                if element.attrib ['id'] == "MaxSeitenZahl":
                    #print("TextElement detected in existing invoice element tree: " + str(element.attrib['id']))
                    element.attrib['id'] = "MaxSeitenZahl" + str(p-2)
                
                for i in range(0,p):
                    element = newtree.xpath('//*[@id="MaxSeitenZahl%s"]' % str(i))
                    
                    for spans in element:
                        #print("Children of TextElement detected: " + str(span.attrib['id']))

                        for span in spans:
                            for child in span:
                                #print("Children of Child detected: " + str(child.attrib['id']))
                                child.text = str(p-1)
                                #print("replaced MaxSeitenZahl to: "+ str(p-1))


            #here i open the original template file to get the article row with placeholders as a template to fill in again
            f = open(templateFile,'r')
            filedata = f.read()
            f.close()

            #get fresh element tree from original template svg file
            parser=etree.XMLParser(encoding='utf-8')
            freshtree = etree.parse(templateFile, parser)
            root = freshtree.getroot()

            #probably easier to get but this works for now
            for element in freshtree.xpath('//*[local-name()="svg"]//*[local-name()="g"]', namespaces=SVGNS):

                #basically same as above, but here we want the fresh element instead the placement in the tree
                if element.attrib['id'] == "ArticleRow0":
                    #print("ArticleRow detected in unchanged template element tree: " + str(element.attrib['id']))
                    nRow = element
                    parent = nRow.getparent()
                    #print("found parent with ID: " + parent.attrib['id'])
                    #print()

            # here we add the template article row to the last saved invoice and change some attributes, like id and position in file       
            nRow.attrib['id'] = 'ArticleRow' + str(n-1)
            #print("Changed nRow name to: " + str(nRow.attrib['id']))
            nRow.attrib['transform'] = "translate(0," + str(translation) +")"
            pos.addnext(nRow)         
            print("added ArticleRow" + str(n-1))
            print()
            
            #this converts the full modified svg tree element to a string, which we then safe
            newdata = etree.tostring( newtree, pretty_print=True, xml_declaration=True, standalone="yes").decode()

            #add next spacing
            translation = translation + spacing
            #print("next translation: " + str(translation))

        #this is a modfier for a new page after 7 elements, afterwards goes back to n>1 loop
        if n % 8 == 0:
            print("n größer 7 - neue Seite")
            print("Seite: " + str(p-1))

            translation = 60 # this is the amount of space the next article Row needs to be offsetet below the previous one. we add them app during the loop to stack the rows
            #open the original invoice template with placeholders to generate page two
            f = open(templateFile, 'r')
            filedata = f.read()
            f.close()            

            #get fresh element tree from original template svg file
            parser=etree.XMLParser(encoding='utf-8')
            pagetree = etree.parse(templateFile, parser)
            root = pagetree.getroot()

            #probably easier to get done, but this works for now
            for element in pagetree.xpath('//*[local-name()="svg"]//*[local-name()="page"]', namespaces=SVGNS):

                #grabbing the page element
                if element.attrib['id'] == "Page0":
                    #print("Element detected " + str(element.attrib['id']))
                    page = element
                    page.attrib['id'] = 'Page' + str(p-1)
                    page.attrib['x'] = str(translationX)

            for element in pagetree.xpath('//*[local-name()="svg"]//*[local-name()="g"]', namespaces=SVGNS):

                #grabbing the full page element
                if element.attrib['id'] == "contentGroup0":
                    #print("Element detected " + str(element.attrib['id']))
                    content = element
                    contentParent = content.getparent()
                    #print("parent element: ")
                    #print(contentParent)
                    content.attrib['id'] = "contentGroup" + str(p-1)
                    content.attrib['transform'] = "translate(" + str(translationX) + ")"

                    for group in content:                     
                        
                        if group.attrib['id'] == "Summen":
                            group.attrib['id'] = "Summen" + str(p-1)
                            #print("Renamed Summen of contentgroup to: " + str(group.attrib['id']) )

                        if group.attrib['id'] == "tabelle":
                            group.attrib['id'] = "tabelle" + str(p-1)
                            #print("Renamed Tabelle of contentgroup to: " + str(group.attrib['id']) )

                            for articlerow in group:

                                if articlerow.attrib['id'] == "ArticleRow0":
                                    articlerow.attrib['id'] = "ArticleRow" + str(n-1)
                                    #print("Renamed ArticleRow0 of contentgroup to: " + str(articlerow.attrib['id']))

            #now we want to insert it to our presaved temporary customer invoice
            f = open(invoiceFile, 'r')
            filedata = f.read()
            f.close()

            parser=etree.XMLParser(encoding='utf-8')
            oldtree = etree.parse(invoiceFile, parser)
            root = oldtree.getroot()

            for element in oldtree.xpath('//*[local-name()="svg"]//*[local-name()="page"]', namespaces=SVGNS):

                #grabbing the full page element
                if element.attrib['id'] == "Page" + str(p-2):
                    #print("Element detected " + str(element.attrib['id']))
                    pagePosition = element
                    pagePositionParent = pagePosition.getparent()
                    #print("parent element: ")
                    #print(pagePositionParent)
                    #pagePosition.attrib['id'] = 'Page' + str(p-2)


            for element in oldtree.xpath('//*[local-name()="svg"]//*[local-name()="g"]', namespaces=SVGNS):

                #grabbing the full content group element with all the elements
                if element.attrib['id'] == "contentGroup" + str(p-2):
                    #print("Element detected " + str(element.attrib['id']))
                    contentPosition = element
                    contentPositionParent = content.getparent()
                    #print("parent element: ")
                    #print(contentParent)
                    content.attrib['id'] = "contentGroup" + str(p-1)
                    content.attrib['transform'] = "translate(" + str(translationX) + ")"
                
                if element.attrib['id'] == "HinweisZwischenSumme":
                    hinweis = element
                    #print("Element detected " + str(element.attrib['id']))
                    element.attrib['id'] = "HinweisZwischenSumme" + str(p-2) #translate(0,415)
                    element.attrib['transform'] = "translate(0,415)"

                    hinweisParent = hinweis.getparent()
                    #print("Parent Element: " + str(hinweisParent))
                    SummenParent = hinweisParent.getparent()
                    #print("Parent Element: " + str(SummenParent))
                
                    for text in hinweis:
                        if text.attrib['id'] == "ZwischenSumme":
                            #print("Textelement in Hinweis detected " + str(text.attrib['id']))
                            text.attrib['id'] = "ZwischenSumme" + str(p-2)

                    SummenParent.append(hinweis)                    
                    SummenParent.remove(hinweisParent)
                    
            pagePosition.addnext(page)                        
            print("added new Page: " + str(page.attrib['id']))
            print()
            contentPosition.addnext(content)
            #print("added new Content: " + str(content.attrib['id']))
            print()

            #this converts the full svg tree element to a string, which we then safe
            newdata = etree.tostring( oldtree, pretty_print=True, xml_declaration=True, standalone="yes").decode()  
            # Fill in Adress
            newdata = newdata.replace(platzhalterName,row[17])
            print("replaced Name")
            newdata = newdata.replace(platzhalterStrasse,row[18])
            print("replaced Strasse")
            newdata = newdata.replace(platzhalterPLZ,row[22])
            print("replaced PLZ")
            newdata = newdata.replace(platzhalterStadt,row[20])
            print("replaced Stadt")

            #Fill in Order Dates
            newdata = newdata.replace(platzhalterBestellDatum,row[0])
            print("replaced BestellDatum")
            newdata = newdata.replace(platzhalterBestellNr,row[24])
            print("replaced BestellNr")
            newdata = newdata.replace(platzhalterRechnungsDatum,row[0])
            print("replaced Rechnungsdatum")
            newdata = newdata.replace(platzhalterRechnungsNr,row[24])
            print("replaced RechnungsNr")

            #Fill in Article Stuff that is static
            newdata = newdata.replace(platzhalterVersandkosten,row[9])
            print("replaced Versandkosten zu: "+ str(format(float(row[9]), '.2f')))   

            p=p+1 
            translationX = translationX + spacingX   

        #Fill in Article Stuff that is dynamic - if someone only has one article
        newdata = newdata.replace(platzhalterArtikelNr,row[32])    
        print("replaced ArtikelNr zu: "+ str(row[32])) 
        newdata = newdata.replace(platzhalterArtikelName,row[1])
        print("replaced ArtikelName zu: "+ str(row[1]))
        newdata = newdata.replace(platzhalterAnzahl,row[3])
        print("replaced ArtikelAnzahl zu: "+ str(row[3]))
        newdata = newdata.replace(platzhalterPreis,format(Preis, '.2f')) 
        print("replaced ArtikelPreis zu: "+ str(format(Preis, '.2f')))
        newdata = newdata.replace(platzhalterGesamtPreis,format(GesamtPreis, '.2f')) 
        print("replaced GesamtPreis zu: "+ str(format(GesamtPreis, '.2f')))    
        newdata = newdata.replace(platzhalterZwischenSumme,format(ZwischenSumme, '.2f')) 
        print("replaced ZwischenSumme zu: "+ str(format(ZwischenSumme, '.2f')))  
        newdata = newdata.replace(platzhalterMWST,format(MWST, '.2f')) 
        print("replaced MWST zu: "+ str(format(MWST, '.2f')))    

        newdata = newdata.replace(platzhalterSeitenZahl,"Seite " + str(p-1)) 
        print("replaced Seitenzahl zu: "+ "Seite " + str(p-1))
        newdata = newdata.replace(platzhalterMaxSeitenZahl, str(p-1)) 
        print("replaced MaxSeitenzahl zu: "+ str(p-1))        

        f = open(invoiceFile,'w')
        f.write(newdata)
        f.close()

        #convert to pdf with inkscape - filename is with order id and customer name
        invoicePDF = filepath + r'\invoice\pdf\etsy-invoice-%(id)s-%(name)s.pdf' % { 'id' : str(row[24]), 'name' : str(row[17])}
        print("converting to pdf: " + invoicePDF)

        completed = subprocess.run(['C:/Program Files/Inkscape/bin/inkscapecom.com',  invoiceFile, r'--export-filename=' + invoicePDF])

        print ("stderr:" + str(completed.stderr))
        print ("stdout:" + str(completed.stdout))
        print()

        #print("Durchgang: " + str(n))
        n=n+1
        print()
        previousCustomer = row[17]
        #print("letzter Kunde: " + str(previousCustomer))
        print()
        print()

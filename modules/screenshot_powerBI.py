# -*- coding: utf-8 -*-
"""
Created on Tue Oct 15 10:35:09 2019

@author: Felix
"""

#%% IMPORT LIBRARIES

print("Lesewert Screenshot-Modul geladen!")

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
import seaborn as sns
import datetime as dt

# Deutsch als mathematischen Standard einsetzen
import locale
locale.setlocale(locale.LC_ALL, "deu_deu")


# import dt. Mathezeichen (Komma statt punkt bei Dezimalzahlen)
matplotlib.rcParams['axes.formatter.use_locale'] = True


#import der Schrift für Matplotlib
from matplotlib import rcParams
import sys
import os
import matplotlib.font_manager as fm
#Schrift laden
fpath1 = 'C:\\Windows\\Fonts\\Campton-Light.otf'
campton_light = fm.FontProperties(fname=fpath1)


import timeit
import time
import requests
#für die Bildbearbeitung:
from PIL import Image
from PIL import ImageDraw, ImageFont
from io import StringIO
import textwrap




#%% Screenshot-Funktion

def screenshot_powerbi(df, ressort = "Ressort", folder="screenshots", number=5, rangliste=False, ausgabe = "", kunden_id =  0, 
               zeitung=True, mode="ressort", outside=False, markengröße="mittel"):
    LW = "Artikel-Lesewert (Erscheinung) in %"
    BW = "Artikel-Blickwert (Erscheinung) in %"
    DW = "Artikel-Durchlesewerte (Erscheinung) in %"
    seite = "Seitennummer"
    LW = "LW"
    BW = "BW"
    DW = "DW"
    id_nr = kunden_id
    
    df["Seitennummer"] = df["Seite"]
    seite = "Seite"
    
    
        
    df = df.sort_values(by=LW, ascending=False).head(number)
    for i in range(df.shape[0]):
        #id_nr = "1011" #für die Schwäbische Zeitung
        size_marke = markengröße
        
        elem = df.iloc[i]
       
        # normale URL zum Laden der Screenshots
        url = "https://lesewert.azureedge.net/layer/"+str(id_nr)+"/"+ str(elem.AusgabeId)+ "/"+ str(elem.ErscheinungsId) + "/"+\
        str(elem.Seitennummer) + "/" +  str(elem.ArtikelId) +  ".jpg"
        # Ausweich-URL zum Laden der Screenshots
        url2= "https://mein.lesewert.de/Offline/ContentProxy?url=" + str(id_nr) + "/"+ str(elem.AusgabeId) + "/"+ \
        str(elem.ErscheinungsId) + "/"+ str(elem[seite]) + "/" +  str(elem.ArtikelId) +  ".jpg"
        
        response = requests.get(url, stream=True)
        response.raw.decode_content = True
    
        # Try-Block prüft, ob URL antwortet
        # Wenn nicht wird alternative URL geladen
        try:
            im = Image.open(response.raw)
        except OSError:
            print("ACHTUNG FEHLER")
            print("Konnte folgende URL nicht laden:") 
            print(url)
                       
            response2 = requests.get(url2, stream=True)
            response2.raw.decode_content = True
            im = Image.open(response2.raw)
    
        #Breite des Bildes verringern, wenn der Text höher als breit ist, z.B. ganze Reportage-Seite
        # nur dann ist garantiert, dass die ÜS im Bild ist. 
        # pciture_seize ist var für späteren Bildausschnitt, so dass kein Weißgraum entsteht
        picture_seize = (1000,377)
        if im.size[0] < im.size[1]:
            corr = im.size[1]/im.size[0]
            if ausgabe == "schwaebische":
                height = int(400*corr)
                im = im.resize((400, height), Image.ANTIALIAS)
            else:
                height = int(550*corr)
                im = im.resize((550, height), Image.ANTIALIAS)
                picture_seize = (550, height)
                if im.size[1] / im.size[0] >= 1.7:
                    size_marke = "groß"
        else: 
            corr = im.size[1]/im.size[0]
            if ausgabe == "schwaebische": 
                height = int(600*corr)
                im = im.resize((600, height), Image.ANTIALIAS)
        #TODO:_ Achtung, an den Fotos rumgefummelt
        print(size_marke)
        # Bildbearbeitung mit PILLOW
        # Beschriften der LW-Marke
        marke = Image.open("./lw_ressources/lw_marke_png.png")
        draw = ImageDraw.Draw(marke)
        
        font = ImageFont.truetype('Campton-Light.otf', size=60) 
        font_bold = ImageFont.truetype('Campton-Bold.otf', size=60) 
        (x1,y1) = (340,200)
        (x2,y2) = (340, 300)
        (x3,y3) = (340, 400)
        (x4,y4) = (340, 490)
        #message1 = str(elem["Platzierung"])
        
        message1 = str(elem["Platzierung"]) + str(".")
        message2 = str(round(elem[LW],1)).replace(".", ",") + "%"
        message3 = str(round(elem[BW],1)).replace(".", ",") + "%"
        message4 = str(round(elem[DW],1)).replace(".", ",") + "%"
        color="rgb(255, 255, 255)"
        draw.text((x1,y1), message1, fill=color, font=font)
        draw.text((x2,y2), message2, fill=color, font=font_bold)
        draw.text((x3,y3), message3, fill=color, font=font)
        draw.text((x4,y4), message4, fill=color, font=font)
    
        # Größe der Marke anpassen
        if size_marke == "klein":
            marke = marke.resize((629, 472), Image.ANTIALIAS)
        elif size_marke == "groß": 
            marke = marke.resize((1625, 1250), Image.ANTIALIAS)
        else: 
            marke = marke.resize((944, 708), Image.ANTIALIAS)
            
            
        
           
        
        #Zwischenspeichern der Marke
        # Normale Größe = (1000, 377), ist unter picture_size gespeichert
        final1 = Image.new("RGBA", picture_seize, (255,255,255,255))
        # TODO Feineinstellung, Ausgang (0x,0y)
        final1.paste(im, (0,0))
        
        if size_marke == "klein":
            
            final1.paste(marke, (20, 150), marke)
        elif size_marke == "groß": 
            final1.paste(marke, (20, 100),marke)
                
        else:
            if im.size[0]<im.size[1]:        # ist der Artikel lang, wird er kleiner, Marke rutscht nach rechts oben  
                final1.paste(marke, (200,150), marke) # 650 
            else: # ansonsten bleibt die Marke unten rechts
                final1.paste(marke, (530,35), marke) # 530
        
            
        print(f"Bild {elem[ressort]} {i} hat die Maße {im.size[0]} zu {im.size[1]}.")    
        filename = "C:/Users/Felix/Desktop/LW/"+ folder + "/" + elem[ressort].replace(" ", "").replace("/", "_") +"_" +  str(i) +".png"
        print(filename)
        final1.save(filename)  
    
#%%
        
   #mode = bezeichnet die Kategorie, die durchlaufen wird (Standard Ressort, aber auch Seitentitel etc. möglich)      


# Helper Function - Screenshot ausspielen 
def create_screenshot(df_, id_nr, seite, ausgabe, LW, BW, DW, different_size, folder, ressort ): 
    for i in range(df_.shape[0]):
            #id_nr = "1011" #für die Schwäbische Zeitung
            
            
            el = df_.iloc[i]
           
            # normale URL zum Laden der Screenshots
            url = "https://lesewert.azureedge.net/layer/"+str(id_nr)+"/"+ str(el.AusgabeId)+ "/"+ str(el.ErscheinungsId) + "/"+\
            str(el.Seite) + "/" +  str(el.ArtikelId) +  ".jpg"
            
            # Ausweich-URL zum Laden der Screenshots
            url2= "https://mein.lesewert.de/Offline/ContentProxy?url=" + str(id_nr) + "/"+ str(el.AusgabeId) + "/"+ \
            str(el.ErscheinungsId) + "/"+ str(el[seite]) + "/" +  str(el.ArtikelId) +  ".jpg"
            
            response = requests.get(url, stream=True)
            response.raw.decode_content = True
        
            # Try-Block prüft, ob URL antwortet
            # Wenn nicht wird alternative URL geladen
            try:
                im = Image.open(response.raw)
            except OSError:
                
                           
                response2 = requests.get(url2, stream=True)
                response2.raw.decode_content = True
                
                im = Image.open(response2.raw)
        
            #Breite des Bildes verringern, wenn der Text höher als breit ist, z.B. ganze Reportage-Seite
            # nur dann ist garantiert, dass die ÜS im Bild ist. 
            
            # unterschiedliche Einstellungen für different_size (also die Frage,
            # Bildbreite, Markenposition etc. für breitere Texte anders laufen sollen)
            
            if different_size: 
            
                if (im.size[0] < im.size[1]):
                    
                    corr = im.size[1]/im.size[0]
                    if ausgabe == "schwaebische":
                        height = int(400*corr)
                        im = im.resize((400, height), Image.ANTIALIAS)
                    else:
                        height = int(550*corr)
                        im = im.resize((550, height), Image.ANTIALIAS)
                else: 
                    corr = im.size[1]/im.size[0]
                    if ausgabe == "schwaebische": 
                        height = int(600*corr)
                        im = im.resize((600, height), Image.ANTIALIAS)
                
            else:
                
                
                corr = im.size[1]/im.size[0]
                height = int(550*corr)
                width = im.size[0]
                width_for_image = 550
                #wenn Breite unter 120 (also sehr schmale Meldung)_
                # Breite verdoppeln, Höhe anpassen
                if width < 120:
                    width_for_image = int(width * 1.5)
                    height = int(width_for_image * corr )
                    
                elif im.size[0] < 400:
                    height = int(im.size[0]*corr)
                    width_for_image = width
                    
                else: 
                    
                    if ausgabe == "schwaebische":
                        height = int(600*corr)
                        im = im.resize((600, height), Image.ANTIALIAS)
                
                # wenn Artikel besonders flach sind, wird die Breite erhöht
                if height < 250: 
                    width_for_image = 650
                    height = int(650*corr)
                #print(f"Widthforimage: {width_for_image}")
                im = im.resize((width_for_image, height), Image.ANTIALIAS)
                
            # Bildbearbeitung mit PILLOW
            # Beschriften der LW-Marke
            marke = Image.open("lw_marke_png.png")
            draw = ImageDraw.Draw(marke)
            
            font = ImageFont.truetype('Campton-Light.otf', size=60) 
            font_bold = ImageFont.truetype('Campton-Bold.otf', size=60) 
            (x1,y1) = (340,200)
            (x2,y2) = (340, 300)
            (x3,y3) = (340, 400)
            (x4,y4) = (340, 490)
            #message1 = str(elem["Platzierung"])
            
            message1 = str(el["Platzierung"]) + str(".")
            message2 = str(round(el[LW],1)).replace(".", ",") + "%"
            message3 = str(round(el[BW],1)).replace(".", ",") + "%"
            message4 = str(round(el[DW],1)).replace(".", ",") + "%"
            color="rgb(255, 255, 255)"
            draw.text((x1,y1), message1, fill=color, font=font)
            draw.text((x2,y2), message2, fill=color, font=font_bold)
            draw.text((x3,y3), message3, fill=color, font=font)
            draw.text((x4,y4), message4, fill=color, font=font)
       
            # Größe der Marke anpassen
            marke = marke.resize((944, 708), Image.ANTIALIAS)
                
            
            #Zwischenspeichern der Marke
            if different_size:
                height_white = 377
                final1 = Image.new("RGBA", (1000, height_white), (255,255,255,255))
            else:
                height_white = 450
                final1 = Image.new("RGBA", (1000, height_white), (255,255,255,255))
            # TODO Feineinstellung, Ausgang (0x,0y)
            
            if different_size: 
                position = (0,0)
                final1.paste(im, position)
            else: 
                
                y_axis = (height_white - height) / 2
                if y_axis > 0:
                    y_axis = int(y_axis)
                else: 
                    y_axis = 0
                #print(el[ressort])
                
                #print("Breite: ", int(im.size[0]))
                #print("Höhe: ", height)
                #print("Bildhöhe ", height_white)
                #print("y_axis ", y_axis)
                position= (0, y_axis)
                #print("Position ", position)
                #print("--------------")
                final1.paste(im, position)
            #Unterschiedliche Eintellungen für different_size
            if different_size: 
                if im.size[0]<im.size[1]:        # ist der Artikel lang, wird er kleiner, Marke rutscht nach rechts oben  
                    final1.paste(marke, (650,5), marke)
                else: # ansonsten bleibt die Marke unten rechts
                    final1.paste(marke, (530,35), marke)
            else:
                final1.paste(marke, (650,5), marke)
                
    
            final1.save("final1.png")  
        
            print("Hallo")  
            filename = "C:/Users/Felix/Desktop/LW/"+ folder + "/" + el[ressort].replace(" ", "").replace("/", "_") +"_" +  str(i) + str(el[ausgabe]) +".png"
            print(filename)
            final1.save(filename)  
        
#%%

def screenshot_top10(df, number=5, ressort="Ressort", folder="screenshots_top10", rangliste=False, ausgabe = "", 
               zeitung=True, mode="Ressort", outside=False, kunden_id = 0, different_size=True, Lokales=True):
#    LW = "Artikel-Lesewert (Erscheinung) in %"
#    BW = "Artikel-Blickwert (Erscheinung) in %"
#    DW = "Artikel-Durchlesewerte (Erscheinung) in %"
#    seite = "Seitennummer"
#    if df["LW"]:
    print(f'Wert für Different_size: {different_size}')
    if different_size: 
        print("Different size ist True.")
    else: 
        print("Differentsize ist False.")
        
    LW = "LW"
    BW = "BW"
    DW = "DW"
    seite = "Seite"
    
    ausgabe = "Ausgabe"
    id_nr = kunden_id
    
    ausgabename = "Ausgabename" # sonst auch gerne ZTG
    # Liste aller Ressorts anlegen, nan ausschließen
    list_ = df[mode].unique()
    if kunden_id == 1015:
        list_ = ['Lokales', 'Sport', 'Ostfriesland', 'Lokalsport', 'Wirtschaft',
       'Kinderseite', 'Kultur', 'Hintergrund', 'Wochenende', 'Panorama',
       'Sonderseiten', 'Meinung', 'Thema', 'Warentest / Ratgeber',
       'Politik', 'Titelseite', 'Nordwest']
    list_clean = [x for x in list_ if str(x) != 'nan']
    
    for elem in list_clean: 
        df_ = df[df[mode]==elem]
        
       
        if elem == "Lokales": 
           
            ausg = df_[ausgabe].unique()
           
            for aus in ausg: 
                
                df_lok = df_[df_[ausgabe]==aus]
                df_lok = df_lok.sort_values(by=LW, ascending=False).head(number)
                
                create_screenshot(df_lok, id_nr, seite, ausgabe, LW, BW, DW, different_size, folder, ressort)
        else:
            df_ = df_.sort_values(by=LW, ascending=False).head(number)
            create_screenshot(df_, id_nr, seite, ausgabe, LW, BW, DW, different_size, folder, ressort )
            
   
    

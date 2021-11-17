import pandas as pd
import numpy as np
# time für die Bearbeitungszeit
import time
import locale
locale.setlocale(locale.LC_ALL, "deu_deu")
import os
import shutil
import requests

#Schrift laden
import matplotlib.font_manager as fm
fpath1 = 'C:\\Windows\\Fonts\\Campton-Light.otf'
campton_light = fm.FontProperties(fname=fpath1)
#für die Bildbearbeitung und Beschriftung
from PIL import Image
from PIL import ImageDraw, ImageFont
from io import StringIO, BytesIO
import io
import textwrap
# Adresse zum Speicherort auf einem Server, damit ich die Datei nicht hin- 
# und herkopieren muss. 
url_lesewertmarke = "https://datasets.app.dmwm.de/LW_Marke/lw_marke_png.png"

class Screenshot():
    '''
    Diese Klasse erstellt ein Screenshot-Dataframe. 
    Das Dataframe trägt sämtliche Artikel, die es auszuwerten gilt. 
    Auf dieser Gesamtzahl von Artikeln werden dann die unterschiedlichen 
    Screenshot-Methoden angefertigt
    '''
    
    def __init__ (self, dataframe, kunden_id, \
                  ressort = "Ressort", seitentitel = "Seitentitel", 
                  edition="Ausgabenteil"):
        ''' Method for initializing a Screenshot-Objekt
        
        Args:
        dataframe = Dataframe mit kumulierten Daten
        kunden_id = vierstellig Kunden-Id, von Lesewert vergeben
        
        default Args: 
        ressorts = String, der die Column benennt, in der die Ressorts 
        gespeichert sind.
        editions_local = Column, in der die Ausgabenteile verzeichnet sind 
        ressort: Column-Name, unter dem die Ressorts des DF zu finden sind
        seitentitel: Column-Name, unter dem die Seitentitel zu finden sind
        edition: Columen-Name, unter dem die Ausgabenteile zu finden sind
        
        Attributes:
        dataframe = fertig vorbereitetes Dataframe mit allen Columns (kein 
        Merge mehr nötig)
        kunden_id = vierstellig Kunden-Id, von Lesewert vergeben
        ressor_col: Column-Name, unter dem die Ressorts des DF zu finden sind
        seitentitel_col: Column-Name, unter dem die Seitentitel zu finden sind
        ausgaben_col: Columen-Name, unter dem die Ausgabenteile zu finden sind
        
        
        edition = alle Ausgabenteile (inklusive Mantel, Sonntag u.ä.)
        seitentitel = alle Seitentitel des DF, besonders interessant, wenn ich
            als dataframe nur ein Ressort übergebe
        
        '''
        self._dataframe = dataframe
        self._kunden_id = kunden_id
        self._ressort_col = ressort
        self._seitentitel_col = seitentitel
        self._ausgaben_col = edition
        
        # attributes defined inside class
        self._ressorts = self._dataframe[self._ressort_col].unique()        
        self._original_ressorts = self._ressorts
        self._editions = self._dataframe[self._ausgaben_col].unique()
        self._seitentitel = self._dataframe[self._seitentitel_col].unique()

        self._lokale_ressorts = ["Lokales", "Lokalsport", "Lokale Kultur",
            "Regionalsport", "Kultur lokal", "Kultur regional"]
        

        
    
    # Getter
    @property
    def dataframe(self):
        return self._dataframe
    @property
    def kunden_id(self):
        return self._kunden_id
    @property
    def ressort_col(self):
        return self._ressort_col
    @property
    def seitentitel_col(self):
        return self._seitentitel_col
    @property
    def ausgaben_col(self):
        return self._ausgaben_col
    @property
    def ressorts(self):
        return self._ressorts
    @property
    def original_ressorts(self):
        return self._original_ressorts
    @property
    def editions(self):
        return self._editions
    @property
    def seitentitel(self):
        return self._seitentitel
  
    # Setter 
    @dataframe.setter
    def dataframe(self, dataframe):
        self._dataframe = dataframe
    @kunden_id.setter
    def kunden_id(self, kunden_id):
        self._kunden_id = kunden_id
    @ressort_col.setter
    def ressort_col(self, ressort_col):
        self._ressort_col = ressort_col
    @seitentitel_col.setter
    def seitentitel_col(self, seitentitel_col):
        self._seitentitel_col = seitentitel_col
    @ausgaben_col.setter
    def ausgaben_col(self, ausgaben_col):
        self._ausgaben_col = ausgaben_col
    @ressorts.setter
    def ressorts(self, ressorts):
        self._ressorts = ressorts
    @original_ressorts.setter
    def original_ressorts(self, original_ressorts):
        raise Exception("Achtung, Änderung der Original-Ressorts nicht \
            möglich, bitte 'ressorts' ändern.")
    @editions.setter
    def editions(self, editions):
        self._editions = editions
    @seitentitel.setter
    def seitentitel(self, seitentitel):
        self._seitentitel = seitentitel
   

    ### HELPER-Methods
    def make_screenshot_dir(self, root_dir = "screenshots",\
        delete_existing_folder=True):
        ''' 
        erstellt eine Directory, in der die Screenshots abgelegt werden 
        löscht bei Aufruf bereits vorhandenen Ordner
        '''
        if os.path.exists(root_dir) and delete_existing_folder:
            shutil.rmtree(root_dir)
            working_dir = os.getcwd()
            print(f"Vorhandender Ordner {working_dir}\\{root_dir}\
                 mit allen Unterordnern gelöscht.")

        try:
            os.makedirs(root_dir)
        except OSError:
            print (f"Ordner {root_dir} konnte nicht erstellt werden")
        else:
            print (f"Ordner {root_dir} erfolgreich erstellt")


    
    def _making_screenshot(self, df, part_of_filename, folder="screenshots", \
        size_marke = "mittel", number_screenshots = 5, different_size=False, \
            mode="Ressort", gallery = False): 
        '''
        Function sortiert das Dataframe nach LW
        fertigt dann einen Screenshot an
        produziert die Lesewertmarke und legt sie über den Screenshot
        spielt die Datei im Ordner screenshots aus

        Args: 
        df = hier benötigen wir ein finales DF, bereinigt und mit allen 
            Columns gemerged, und ausschließlich mit Artikel aus einer 
            Auswertungseinheit (z.B. einem Ressort)
        part_of_filename = trägt den späteren Filenamen (der plus den Iterator 
            abgespeichert wird) 
        folder - trägt den Namen des Überordners, in den alle Screenshots 
            abgelegt werden, normalerweise "screenshots"
        size_marke - kann man notfalls händisch festlegen, ist aber in der 
            Regel nicht notwendig
        number_screenshot - legt die Anzahl der Screenshots pro 
            Auswertungseinheit (mode) fest
        different_size = zeigt an, wie sich Größe und Position der Marke etc 
            verändern. 
            Voreinstellung True, die Marke setzt sich dann in unterschiedlicher
            Größe auch mal auf die Artikel. 
            Das wird benötigt, um die Gallery-Bilder gut nebeneinander 
            platzieren zu können. 
            False: Steht immer in derselben Größe neben den Artikeln. 
        mode - hier können wir die Screenshots nach Ressort oder nach \
            Seitentiteln anfertigen lassen, 
            voreingestellt ist Ressort
        gallery = zeigt an, ob wir eine Gallery für Analysezwecke (z.B. die 50
            besten Artikel) anlegen = True, 
            oder ob wir False eine Top-10 Ansicht planen 
            (ähnlich wie different_size).
            Außerdem entscheidet gallery über die Ornderstrukur: 
            FALSE = screenshots/top_article/{Ressort}
            TRUE =  screenshots/gallery/{Ressort}
        '''
        
        # prüfen, ob die Column-Names noch nicht formatiert wurden
        if "Artikel-Lesewert (Erscheinung) in %" in self._dataframe.columns:
            LW = "Artikel-Lesewert (Erscheinung) in %"
            BW = "Artikel-Blickwert (Erscheinung) in %"
            DW = "Artikel-Durchlesewerte (Erscheinung) in %"
            seite = "Seitennummer"
        else: 
            LW = "LW"
            BW = "BW"
            DW = "DW"
            seite = "Seite"
       
        #Dataframe absteigend sortieren, höchster LW zuerst
        df = df.sort_values(by=LW, ascending=False).head(number_screenshots)
        
        #Loop über jeden Artikel des geordneten und begrenzten df-Dataframes
        for i in range(df.shape[0]):
            platzierung = i+1
            #var für die jeweilige line
            df_line = df.iloc[i]
            # Bild-URL für jeweiligen Artikel
            url1 = "https://lesewert.azureedge.net/layer/"\
                +str(self._kunden_id)+"/"+ str(df_line.AusgabeId.astype(int))\
                    + "/"+ str(df_line.ErscheinungsId.astype(int)) + "/"\
                        +str(df_line[seite]) + "/" +  str(df_line.ArtikelId) \
                            +  ".jpg"
            url2 = "https://mein.lesewert.de/Offline/ContentProxy?url=" \
                + str(self._kunden_id) + "/"\
                    + str(df_line.AusgabeId.astype(int)) + "/"\
                        + str(df_line.ErscheinungsId.astype(int)) + "/"\
                            + str(df_line[seite]) + "/" \
                                +  str(df_line.ArtikelId) +  ".jpg"



            response = requests.get(url1, stream=True)
            response.raw.decode_content = True
           
            # prüfen, ob sich die response als Image öffnen lässt, 
            # wenn nicht: alternative URL2 benutzen
            try:
                im = Image.open(response.raw)

            #except OSError:
            except:
                response2 = requests.get(url2, stream=True)
                response2.raw.decode_content = True
                im = Image.open(response2.raw)


            # Bildformat bearbeiten
            # Bildbreite - verringern, wenn Text höher als breit (also lang)
           
            if different_size:
                picture_s = (1000,377)
                if im.size[0] < im.size[1]:
                    #Bildverhältnis errechen
                    corr = im.size[1]/im.size[0]
                    height = int(550*corr)
                    picture_s = (550, height)
                    im = im.resize((550, height), Image.ANTIALIAS)
                    # bei extremen Formaten -> LW-Marke vergrößen
                    if im.size[1] / im.size[0] >= 1.7:
                        size_marke = "groß"          
                else: 
                    corr = im.size[1]/im.size[0]
                    
                    height = int(900*corr)
                    im = im.resize((900, height), Image.ANTIALIAS)
    #                     if ausgabe == "schwaebische": 
    #                         height = int(600*corr)
    #                         im = im.resize((600, height), Image.ANTIALIAS)
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
                    pass
                # wenn Artikel besonders flach sind, wird die Breite erhöht
                if height < 250: 
                    width_for_image = 650
                    height = int(650*corr)
                
                im = im.resize((width_for_image, height), Image.ANTIALIAS)
                
            
            
            # Marke zeichen, Werte in Marke eintragen
          
            marke = Image.open(requests.get(url_lesewertmarke, stream=True).raw)
            draw = ImageDraw.Draw(marke)

            font = ImageFont.truetype('Campton-Light.otf', size=60) 
            font_bold = ImageFont.truetype('Campton-Bold.otf', size=60) 
            (x1,y1) = (340, 200)
            (x2,y2) = (340, 300)
            (x3,y3) = (340, 400)
            (x4,y4) = (340, 490)

            message1 = str(platzierung) + str(".")
            message2 = str(round(df_line[LW],1)).replace(".", ",") + "%"
            message3 = str(round(df_line[BW],1)).replace(".", ",") + "%"
            message4 = str(round(df_line[DW],1)).replace(".", ",") + "%"
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

            # MERGEN der der verschiedenen Bild-Ebenen
            # weißer Hintergrund
            # Größe weißer Hintergrund/Canvas festlegen
             #Zwischenspeichern der Marke
            if different_size:
                #height_white = 377
                final1 = Image.new("RGBA", picture_s, (255,255,255,255))
            else:
                height_white = 450
                final1 = Image.new("RGBA", (1000, height_white), \
                    (255,255,255,255))
            
            # Screenshot auf Weißfläche posten
            # Position des Screenshots auf dem Canvas bestimmen
            if different_size: 
                position = (0,0)
                final1.paste(im, position)
            else: 
                
                y_axis = (height_white - height) / 2
                if y_axis > 0:
                    y_axis = int(y_axis)
                else: 
                    y_axis = 0
               
                position= (0, y_axis)
                
                final1.paste(im, position)
          
            
            # Marke auf Weißfläche und Screenshot positionieren
            if size_marke == "klein":
                final1.paste(marke, (20, 150), marke)
            elif size_marke == "groß": 
                final1.paste(marke, (20, 100),marke)        
            else:
                if different_size:
                    #ist der Artikel lang, wird er kleiner, Marke rutscht nach 
                    # rechts oben  
                    if im.size[0]<im.size[1]:        
                        final1.paste(marke, (650,5), marke)
                    else: # ansonsten bleibt die Marke unten rechts
                        final1.paste(marke, (530,35), marke)
                else:
                    final1.paste(marke, (650,5), marke) # 530, 35
            
            #Festlegung der Ordnerstruktur, Erstellung der Ordner
            name_of_file = f"{part_of_filename.replace(' ', '').replace('/', '_')}_{str(i)}.png"
            if gallery:
                folder = f"./screenshots/gallery/{part_of_filename}/"               
            else:
                folder = f"./screenshots/top_article/{part_of_filename}/"  
            filename = f"{folder}{name_of_file}"
           
            #Unterordner anlegen
            #Ordner wird nur beim ersten Screenshot (i=0) erstellt, 
            # ansonsten würde er immer wieder überschrieben
            if i == 0:         
                self.make_screenshot_dir(root_dir=folder)
                
            # beim letzten Screenshot (i=number-1) wird eine Erfolgsmeldung 
            # ausgeworfen. 
            if i == number_screenshots-1:
                
                print(f"{number_screenshots} Screenshots {part_of_filename} erstellt.")
                print("---------------------------------------------------")
               
            final1.save(filename)

    
    
            
    
    
         
    # Screenshot-Function mit der Hauptfunktion / HELPER-FUNCTION
    # eigene Function verhindert Code-Doppelungen bei Schleifen
    # über Ressorts UND Ausgabenteil bei Ressort = Lokales

    
    
    def _prepare_data_for_screenshots(self, number_screenshots=10, \
        size_marke="mittel" , mode="Ressort", gallery=True, \
            different_size = True):
        '''
        Methode fertigt aus Dataframe Screenshots mit Lesewertmarken an
        
        Kwargs: 
        number_screenshot - legt die Anzahl der Screenshots pro 
        Auswertungseinheit (mode) fest
        size_marke - kann man notfalls händisch festlegen, ist aber in der 
        Regel nicht notwendig
        mode - hier können wir die Screenshots nach Ressort oder nach 
        Seitentiteln anfertigen lassen, 
        voreingestellt ist Ressort
        gallery = True - gallery meint Bilderstrecke für Top-50-Artikel, 
        gallery=False meint
        
        
        
        '''
        
        # VARIABLEN FESTLEGEN 
    
        # prüfen, ob die Column-Names noch nicht formatiert wurden
        
        if "Artikel-Lesewert (Erscheinung) in %" in self._dataframe.columns:
            LW = "Artikel-Lesewert (Erscheinung) in %"
            BW = "Artikel-Blickwert (Erscheinung) in %"
            DW = "Artikel-Durchlesewerte (Erscheinung) in %"
            seite = "Seitennummer"
        else: 
            LW = "LW"
            BW = "BW"
            DW = "DW"
            seite = "Seite"
        
        # richtige Formate verteilen
        # Prüfung, ob LW etwas anderes ist als float-Format, dann umwandeln
        
        if not isinstance(self._dataframe[LW].iloc[0], float):
            self._dataframe[LW] = self._dataframe[LW].str.replace(",",".")\
                .fillna(0).astype(float)
            self._dataframe[BW] = self._dataframe[BW].str.replace(",",".")\
                .fillna(0).astype(float)
            self._dataframe[DW] = self._dataframe[DW].str.replace(",",".")\
                .fillna(0).astype(float)
        # Prüfung, ob erscheinungsId etwas anderes ist als int, dann umwandeln
        if not isinstance(self._dataframe["ErscheinungsId"].iloc[0], int):
            self._dataframe["ErscheinungsId"] = \
                self._dataframe["ErscheinungsId"].fillna(0).astype(int)
            self._dataframe["AusgabeId"] = \
                self._dataframe["AusgabeId"].fillna(0).astype(int)


        id_nr = self._kunden_id
           
        # iterieren über Ressorts und Lokalteile
        # Listen erstellen
        liste_ressorts = self._ressorts
        liste_ausgaben = self._editions
       
        for res in liste_ressorts:
            # bei Ressort Lokales jeweils die einzelen Ausgaben betrachten
            if res == "Lokales":
                df_lok = \
                    self._dataframe[self._dataframe[self._ressort_col]==res]
                for aus in liste_ausgaben:
                    if (aus != "Mantel" and aus != "Sonntag"):
                        df_aus = df_lok[df_lok[self._ausgaben_col]==aus]
                        part_of_filename = aus

                        self._making_screenshot(df_aus, part_of_filename, 
                        size_marke = size_marke, 
                        number_screenshots = number_screenshots,
                        mode=mode, gallery = gallery, different_size=gallery)
            else: 
                df_res = \
                    self._dataframe[self._dataframe[self._ressort_col]==res]
                part_of_filename = res.replace(" ", "_").replace("/", "_")

                self._making_screenshot(df_res, part_of_filename, 
                size_marke = size_marke, 
                number_screenshots = number_screenshots, mode=mode, 
                gallery = gallery, different_size=gallery)
             
            
            
    def top_screenshots(self, number_screenshots=10 ,size_marke="mittel", 
        mode="Ressort", gallery=False ):
        '''
        Funktion, die die Datenverarbeitung für die Top-10-Screenshots startet. 
        Außerdem: Ein Timer, der die Dauer der Screenshoterstellung anzeigt. 
        
        
        Anmerkung: Dieser Schritt ist notwendig, um dem Nutzer zwei 
        verschiedene screenshot-Methoden 
        zur Verfügung zu stellen (top10_screenshots und gallery_screenshots)
        
        Angewählt werden die jweiligen Methoden über gallery FALSE/TRUE. 
        
        
        '''
        print(f"Top-10-Screenhotsgestartet, {number_screenshots} angefordert.")
        start_time = time.time()
        self._prepare_data_for_screenshots(
            number_screenshots=number_screenshots, size_marke=size_marke, 
            mode=mode, gallery=gallery, different_size = gallery)
        end_time = time.time()
        duration = end_time - start_time
        print()
        print(f"Top-10-Screenshots nach {round(duration,2)} \
            Sekunden abgeschlossen.")

    def gallery_screenshots(self, number_screenshots=10 ,size_marke="mittel", 
        mode="Ressort", gallery=True):
        '''
        Funktion, die die Datenverarbeitung für die Gallery-Screenshots startet. 
        Außerdem: Ein Timer, der die Dauer der Screenshoterstellung anzeigt. 
        
        
        Anmerkung: Dieser Schritt ist notwendig, um dem Nutzer 2 verschiedene 
        screenshot-Methoden zur Verfügung zu stellen 
        (top10_screenshots und gallery_screenshots)
        
        Angewählt werden die jweiligen Methoden über gallery FALSE/TRUE. 
        
        
        '''
        print(f"Top-10-Screenhotsgestartet, {number_screenshots} angefordert.")
        start_time = time.time()
        self._prepare_data_for_screenshots(
            number_screenshots=number_screenshots, 
            size_marke=size_marke, mode=mode, gallery=gallery,
            different_size = gallery)
        end_time = time.time()
        duration = end_time - start_time
        print()
        print(f"Gallery-Screenshots nach {round(duration,2)} Sekunden \
            abgeschlossen.")
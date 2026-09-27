# Dynamische databronnen in Excel / Power Query

Met het configuratiewerkblad bepaal je welke bestanden je wilt importeren via Power Query, zonder een vaste bestandslocatie in elke query op te nemen. Je legt eerst de basislocatie vast, stelt daarna per bestand de naam samen en gebruikt tot slot het benoemde bereik[^1] waarmee de informatie in Power Query opgehaald kan worden.

De Power Query functies **LoadCSV** en **LoadXLS** staan al in het Excel-template. De voorbeelden hieronder gebruiken een CSV-bestand en LoadCSV.

## Stap 1: Bestandslocaties instellen

In de eerste tabel staan de basislocaties van je gegevensbestanden. Elke locatie heeft een unieke naam die met **pad_** begint. Deze naam kan via de dropdown lijst in stap 2 gekozen worden, zodoende hoef je en wijziging in een locatie slechts op 1 plek bij te werken.

- **pad_dynamisch**\
  Verwijst naar de locatie van het geopende Excel-bestand. Als het werkboek met de bijbehorende bronbestanden wordt verplaatst, kan het pad daardoor mee veranderen.\
  Dit pad is alleen zichtbaar nadat het Excel bestand een eerste keer is opgeslagen.

- **pad_downloads**\
  Verwijst naar je persoonlijke Downloads-map. Vervang **_username_** in het pad door je eigen Windows-gebruikersnaam, bijv. `C:\Users\Words-of-Code\Downloads\`.

Een extra bestandslocatie toevoegen:
1.	Gebruik een lege regel in de tabel of, indien nodig, voeg een extra rij toe aan de tabel.
2.	Geef de locatie in **kolom A (Referentie)** een unieke naam, bijvoorbeeld **pad_projecten**.
Deze naam gebruik je ook voor het benoemde bereik[^1] in de **kolom B (Bestandslocatie)**.
3.	Vul in **kolom B (Bestandslocatie)** het pad naar de map in.

*Controleer dat het pad naar een bestaande map verwijst en dat de scheiding tussen de map en de bestandsnaam in de uiteindelijke databron klopt.*

> [!NOTE]
> Een benoemd bereik maken:
> Selecteer de cel die benoemd moet worden. Typ de gewenste naam, bijvoorbeeld *csv_projecten*, in het naamvak links van de formulebalk (waar normaal het celadres staat) en druk op Enter. Je kunt dit ook doen via: **Ctrl+F3 → Namen beheren → Nieuw**.

## Stap 2: Databronnen definiëren

Maak voor ieder te importeren bestand een regel in de tweede tabel. Vul de kolommen met blauwe koppen in; kolom J (Databron) bouwt het volledige pad met de bestandsnaam op via de =DATA.SOURCE() formule. Het resultaat is meteen een controle op je instellingen.

- **Referentie**\
  Geef de databron een herkenbare naam, bijvoorbeeld csv_projecten. Deze naam gebruik je ook voor het benoemde bereik[^1] in de kolom J (Databron).
- **Bestandslocatie**\
  Selecteer een pad_ naam. Door de koppeling met de eerste tabel hoeft een wijziging van een locatie maar op 1 plek te worden gedaan.
- **Subfolder**\
  Vul zo nodig de submap in, bijvoorbeeld brondata. Staat het bestand direct in de gekozen map, laat dit veld dan leeg.
- **Bestandsnaam (basis)**\
  Vul het vaste deel van de naam in. Voor `projecten-20260926.csv` is dit projecten; datum en extensie komen uit de andere kolommen.
- **Bestandsextensie**\
  Vul de extensie zonder punt in, bijvoorbeeld `csv` of `xslx`. LoadCSV ondersteunt momenteel alleen deze twee typen.
- **Met datum?**\
  Vul WAAR of 1 in om de standaard datumtoevoeging te gebruiken. In de template is het standaardpatroon emmdd (jaar maand dag), zoals in `projecten-20260926.csv`. Je kunt hier ook een alternatief datumpatroon invullen. Controleer de uitkomst in Databron, zeker bij verschillen in Excel-taalinstelling.
  - In het datumpatroon zijn de volgende karakters toegestaan: `d`, `m`, `y`, `j`, `e`, `-`, en ` ` (spatie).
  - Het karakter `e` is de weergavetaal onafhankelijke variant voor jaar (bijv. `jjjj` of `yyyy`).
- **Datum overschrijven met…**\
  Vul hier een vaste datum of een ander nummer in wanneer je niet de actuele datum wilt gebruiken. Zet dan ook Met datum? aan.
- **REGEX zoekpatroon**\
  (optioneel) Geef het patroon op van het deel van de opgebouwde bestandsnaam dat je wilt aanpassen.
- **REGEX vervangen met**\
  (optioneel) Geef de vervangende tekst op. Leeg laten verwijdert het gevonden deel.
- **Databron**\
  Controleer hier de volledige bestandslocatie en geef de cel een benoemd bereik, bijvoorbeeld csv_projecten. Zet dezelfde naam in kolom A (Referentie) als geheugensteun.

**Voorbeeld van een databron:**

| Instelling | Waarde |
|:-----------|:-------|
| Bestandslocatie | pad_dynamisch |
| Subfolder | brondata |
| Bestandsnaam basis | projecten |
| Extensie | csv |
| Met datum? | = WAAR |

resulteert bijv. in: `https://organisatienaam.sharepoint.com/personal/_username_/Documents/Desktop/brondata/planning-20260926.csv`

De URL hierboven is alleen een voorbeeld. De waarde in jouw `kolom J (Databron)` moet verwijzen naar het bestand dat je daadwerkelijk kunt openen.

## Stap 3: De databron gebruiken in Power Query

Je kunt eerst via **Gegevens → Gegevens ophalen → Uit bestand** een query laten aanmaken, **of via Gegevens → Gegevens ophalen → Uit andere bronnen → Lege query** beginnen.
Open daarna in de Power Query-editor **Start → Geavanceerde editor**. Daar zie en wijzig je de volledige M-code op één plek.

De twee bronstappen
Neem het pad uit het benoemde bereik[^1] en geef het aan de `LoadCSV` functie door. Binnen een let-blok schrijf je de regels zonder `=` aan het begin van de regel:

```PowerQueryM
  Bronbestand = Text.From(Excel.CurrentWorkbook(){[Name="csv_projecten"]}[Content]{0}[Column1]),
  Bron = LoadCSV(Bronbestand, ",")
```

•	Bronbestand
Leest de tekstwaarde uit de eerste cel van het benoemde bereik csv_projecten in het huidige werkboek.
•	Bron
Laat LoadCSV het bestand openen. Het tweede argument is het scheidingsteken voor het betreffende CSV bestand: gebruik "," voor een komma of ";" voor een puntkomma.
Heb je het bestand eerst via de interface geopend? Voeg Bronbestand boven de bestaande bronstap toe en vervang die bronstap door Bron = LoadCSV(Bronbestand, ","). Laat de overige transformaties staan en controleer daarna het voorbeeld; pas vervolgstappen aan als de structuur van de geladen gegevens daarom vraagt.
Bij een lege query kun je direct dit complete voorbeeld in de Geavanceerde editor plakken:

```PowerQueryM
let
   Bronbestand = Text.From(Excel.CurrentWorkbook(){[Name="csv_projecten"]}[Content]{0}[Column1]),
   Bron = LoadCSV(Bronbestand, ",")
in
   Bron
```

Vervang csv_projecten door de naam van jouw benoemde bereik en kies het scheidingsteken van jouw CSV. 
Voor Excel bestanden dien je de LoadXLS-functie te gebruiken en de naam van het werkblad aan te geven.

```PowerQueryM
let
   Bronbestand = Text.From(Excel.CurrentWorkbook(){[Name="xlsx_projecten"]}[Content]{0}[Column1]),
   Bron = LoadXLS(Bronbestand),
   #"Werkbladnaam" = Bron{[Item="NaamWerkblad",Kind="Sheet"]}[Data]
in
     #"Werkbladnaam"
```



[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/H0E727PHE6)

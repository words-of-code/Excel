# Handleiding databronnen gebruiken in Excel / Power Query

Met dit configuratiewerkblad bepaal je welke bestanden Power Query importeert, zonder een vaste bestandslocatie in elke query op te nemen. Je legt eerst de basislocatie vast, stelt daarna per bestand de naam samen en gebruikt tot slot het benoemde bereik waarmee de informatie in Power Query opgehaald kan worden.

De Power Query functies `LoadCSV` en `LoadXLS` staan al in het Excel-template. De voorbeelden hieronder gebruiken een CSV-bestand en LoadCSV.

## Stap 1: Bestandslocaties instellen

In de eerste tabel staan de basislocaties van je gegevensbestanden. Elke locatie heeft een unieke naam die met **pad_** begint. Deze naam kan via de dropdown lijst in stap 2 gekozen worden.
- **pad_dynamisch**   Verwijst naar de locatie van het geopende Excel-bestand. Als het werkboek met de bijbehorende bronbestanden wordt verplaatst, kan het pad daardoor mee veranderen.
Dit pad is alleen zichtbaar nadat het Excel bestand een eerste keer is opgeslagen.
- **pad_downloads**   Verwijst naar je persoonlijke Downloads-map. Vervang _username_ in het pad door je eigen Windows-gebruikersnaam, bijv. `C:\Users\words-of-code\Downloads\`.

Een extra bestandslocatie toevoegen:
1.	Gebruik een lege regel in de tabel of voeg een extra rij toe aan de tabel.
2.	Geef de locatie in `kolom A (Referentie)` een unieke naam, bijvoorbeeld `pad_projecten`.
Deze naam gebruik je ook voor het benoemde bereik* in de `kolom B (Bestandslocatie)`.
3.	Vul in `kolom B (Bestandslocatie)` het pad naar de map in.

*Controleer dat het pad naar een bestaande map verwijst en dat de scheiding tussen de map en de bestandsnaam in de uiteindelijke databron klopt.*

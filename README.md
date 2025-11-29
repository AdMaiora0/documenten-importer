# Document Importer Handleiding

Deze applicatie helpt bij het automatisch sorteren van documenten in mappen op basis van een Excel-lijst.

## Installatie
1. Zorg dat Python geïnstalleerd is.
2. Installeer de benodigdheden:
   ```
   pip install -r requirements.txt
   ```

## Gebruik
1. Start de applicatie:
   ```
   python src/app.py
   ```
2. **Mapping Bestand**: Selecteer het Excel bestand met de koppeling (bijv. Documentnr -> Patientnr).
3. **Kolommen Selecteren**: Kies welke kolom de bestandsnaam bevat (Bron) en welke kolom het cliëntnummer bevat (Doel).
4. **Mappen Selecteren**:
   - **Bronmap**: De map waar alle losse documenten nu in staan.
   - **Doelmap**: De hoofdmap waar de cliëntmappen aangemaakt moeten worden.
5. Klik op **Start Verwerking**.

## Werking
De app leest het Excel bestand regel voor regel.
- Hij zoekt in de bronmap naar een bestand dat overeenkomt met de 'Bron Kolom' (bijv. "1"). Hij herkent automatisch extensies (bijv. "1.pdf" of "1.docx").
- Hij maakt een map aan in de doelmap met de naam uit de 'Doel Kolom' (bijv. "1513").
- Hij verplaatst het bestand naar die nieuwe map.

## Veiligheid
- De app overschrijft geen bestaande bestanden in de doelmap.
- Als een bestand niet gevonden wordt, wordt dit gemeld in het logboek venster.

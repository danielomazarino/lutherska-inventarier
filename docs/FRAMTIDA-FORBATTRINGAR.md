# Framtida förbättringar

Den här filen samlar idéer som kan vara värdefulla senare men som inte behövs för att börja använda inventarieregistret. En idé ska ligga kvar som **parkerad** tills kyrkan uttryckligen bestämmer att den ska utredas eller byggas.

## Fotodokumentation av inventarier

**Status:** Parkerad 2026-09-01. Ingen funktion ska byggas nu.

### Möjligt värde

Ett foto kan göra det enklare att:

- känna igen ett föremål,
- hitta rätt föremål bland flera liknande,
- dokumentera skick och tillbehör,
- ha underlag vid försäkringsärenden eller inventering.

Nyttan behöver vägas mot mer administration, lagringsutrymme och risken att personer eller känsliga uppgifter råkar komma med på bilderna.

### Alternativ A: foto och uppladdning direkt i appen

En framtida version kan erbjuda:

- **Ta foto** på mobil och iPad,
- **Ladda upp bild** på mobil, Mac och PC,
- förhandsvisning innan bilden sparas,
- möjlighet att ersätta eller ta bort bilden,
- automatisk komprimering så att mobilbilder inte blir onödigt stora,
- uppladdning till en kyrkägd bildmapp i OneDrive.

Kamera och vanlig filuppladdning kan använda samma funktion bakom kulisserna. Skillnaden är bara om bilden kommer direkt från kameran eller redan finns på enheten.

OneDrive bör använda en stabil struktur, exempelvis:

```text
Inventarier/Bilder/{inventarie-id}/{markning}-{namn}-{datum}.jpg
```

Inventarie-ID eller märkning ska vara den stabila kopplingen. Namnet gör filen begriplig för människor men kan ändras senare. Appen behöver även spara OneDrive-filens tekniska ID eller sökväg i Excel. För flera bilder per föremål är en separat tabell, exempelvis `Photos`, bättre än många bildkolumner i `Inventory`.

Följande behöver lösas innan alternativet byggs:

- om varje föremål ska ha en huvudbild eller flera bilder,
- vilka användare som får ladda upp, ersätta och ta bort bilder,
- att samma personer har skrivbehörighet till bildmappen i OneDrive,
- hantering av JPEG, PNG och bilder från iPhone/iPad i HEIC-format,
- maximal bildstorlek och automatisk komprimering,
- vad som händer om nätverket bryts mitt under en uppladdning,
- hur appen visar att en bild är sparad,
- regler för bilder där människor, dokument eller personuppgifter syns.

### Alternativ B: manuell hantering av en administratör

Det här alternativet kan användas utan att appen byggs om:

1. Inventeringspersonalen fotograferar föremålet.
2. Bilden skickas till en utsedd administratör tillsammans med föremålets märkning och namn.
3. Administratören kontrollerar vilket föremål bilden tillhör.
4. Bilden får ett enhetligt filnamn och laddas upp till kyrkans bildmapp i OneDrive.

Exempel på filnamn:

```text
SND-001-Tradlosa-mikrofoner-2026-09-01.jpg
SND-001-Tradlosa-mikrofoner-2026-09-01-02.jpg
```

Använd märkning, namn och datum. Lägg till ett löpnummer om flera bilder tas samma dag. Undvik tecknen `" * : < > ? / \ |`, eftersom de inte får användas i OneDrive-filnamn.

Administratören bör inte behöva gissa vilket föremål en bild tillhör. Den som skickar bilden ska alltid ange:

- inventariemärkning,
- föremålets namn,
- fotografens namn,
- fotograferingsdatum,
- eventuell kommentar om skick eller tillbehör.

Det manuella alternativet är enkelt och kräver ingen ny kod. Nackdelen är mer arbete för administratören, risk för olika filnamn och risk att bilder saknar koppling till rätt inventarie.

Viktigt: en bild som bara laddas upp manuellt till OneDrive visas inte automatiskt i inventarieappen. Den fungerar då som ett separat fotoarkiv. För att visa bilden i appen måste en framtida lösning även spara och läsa bildens OneDrive-referens.

### Rekommendation när frågan tas upp igen

Börja med att fråga hur ofta bilder faktiskt behövs och om de måste synas i appen.

- Om det gäller få bilder och ett separat arkiv räcker, använd den manuella rutinen.
- Om många föremål fotograferas eller bilderna måste visas tillsammans med inventarieuppgifterna, bygg direkt kamera- och uppladdningsstöd i appen.

Ta då ett separat beslut om omfattning, behörigheter och bildhantering innan implementationen börjar.
# Lutherska Inventarier

Interaktivt inventarie- och utlåningsregister för Lutherska Missionskyrkan. Appen använder en Excel-arbetsbok i kyrkans SharePoint eller OneDrive som gemensam datakälla via Microsoft Graph.

Den fullständiga svenska instruktionen för konto, OneDrive/SharePoint, Entra och kostnadsfri GitHub Pages-publicering finns i [docs/INSTALLATION-OCH-PILOT.md](docs/INSTALLATION-OCH-PILOT.md).

## Funktioner

- Inventarier med märkning, antal, anteckningar, kategori, ordinarie plats och ansvarig grupp
- Primär och valfri sekundär ansvarig verksamhetsgrupp
- Dynamiska kategorier, grupper och förvaringsplatser
- Utlåning med låntagare, grupp, registrerare, datum och planerad retur
- Tydlig markering av försenade återlämningar
- Direkt läsning och skrivning mot Excel samt automatisk uppdatering varje minut
- Sampleläge i webbläsaren tills Microsoft 365 har anslutits

## Starta lokalt

```powershell
npm install
npm run dev
```

Kontrollera lösningen med:

```powershell
npm run lint
npm run build
```

## Skapa Excel-arbetsboken

En färdig arbetsbok finns i `workbook/Lutherska-Inventarier.xlsx`. Skapa om den med:

```powershell
npm run workbook:create
```

Arbetsboken innehåller fem namngivna tabeller. Ändra inte tabellnamnen eller kolumnordningen:

| Tabell | Kolumner |
| --- | --- |
| `Categories` | Id, Name, Color |
| `Groups` | Id, Name |
| `Locations` | Id, Name |
| `Inventory` | Id, AssetTag, Name, CategoryId, PrimaryGroupId, SecondaryGroupIds, LocationId, Quantity, Notes |
| `Loans` | Id, ItemId, Borrower, BorrowerGroupId, RecordedBy, LentAt, DueAt, ReturnedAt |

Ladda upp arbetsboken till OneDrive för ett kyrkägt funktions- eller administrationskonto. Om kyrkans exakta paket visar sig innehålla SharePoint kan en kyrkägd dokumentyta användas i stället. Alla användare som ska arbeta i appen behöver redigeringsbehörighet till filen.

## Microsoft 365-anslutning

1. Skapa en appregistrering i Microsoft Entra admin center.
2. Välj **Single-page application (SPA)** och lägg till `http://localhost:5173` som redirect URI för lokal testning. Lägg även till den framtida produktionsadressen.
3. Lägg till delegerad Microsoft Graph-behörighet `Files.ReadWrite`. Excel-API:erna stöder inte application permissions.
4. Kopiera **Tenant ID** och **Application (client) ID**.
5. Skapa en redigerbar delningslänk till arbetsboken i OneDrive.
6. Öppna **Inställningar** i appen, fyll i de tre värdena och välj **Logga in och anslut**.

Inga klienthemligheter lagras i appen. MSAL använder den inloggade användarens delegerade behörighet, och varje Graph-anrop kontrolleras även mot användarens faktiska rättighet till arbetsboken.

## Teknisk synkronisering

Delningslänken löses först till arbetsbokens kanoniska `driveId` och `itemId`. Därefter skapas en persistent workbook session. Appen läser och skriver rader i tabellerna och försöker om tillfälliga svar `429`, `503` och `504` med exponentiell väntan. Synkstatus och fel visas alltid i gränssnittet.

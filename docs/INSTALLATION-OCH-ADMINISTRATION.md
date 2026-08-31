# Installation och administration: Lutherska Inventarier

Den här instruktionen utgår från följande upplägg:

- Koden publiceras kostnadsfritt från Daniels privata GitHub-konto.
- Kyrkan behöver inget GitHub-konto och kyrkans användare behöver aldrig besöka GitHub.
- GitHub innehåller endast den publika appkoden. Inventarier, namn och lån lagras inte på GitHub.
- Kyrkans Excel-arbetsbok ligger i kyrkans OneDrive eller, om tjänsten finns, SharePoint.
- Varje användare loggar in med sitt eget kyrkkonto.
- Ingen separat databas, server eller Azure-prenumeration krävs för lösningen.

Guiden antar att GitHub-projektet heter `lutherska-inventarier`. Då blir webbadressen:

`https://danielomazarino.github.io/lutherska-inventarier/`

Om ett annat projektnamn används måste samma namn bytas i alla webbadresser nedan.

## Vad du kan göra innan du får kyrkkontot

1. Skapa ett **publikt** repository på ditt privata GitHub-konto med namnet `lutherska-inventarier`.
2. Lägg in hela innehållet från `C:\dev\church-inventory` i repositoryt.
3. Kontrollera att `.github/workflows/deploy-pages.yml` finns i repositoryt.
4. Öppna repositoryts **Settings > Pages**.
5. Välj **GitHub Actions** under **Build and deployment > Source**.
6. Låt de tre Microsoft 365-variablerna vara tomma tills du har kyrkkontot.
7. Kontrollera fliken **Actions**. Arbetsflödet **Deploy to GitHub Pages** ska slutföras utan fel.
8. Öppna webbadressen. Appen ska visas i **Sampleläge**.

Kyrkans användare behöver inte GitHub-konton. De får endast webbadressen som en vanlig länk.

## Steg 1: kontrollera kyrkkontot och licensen

När du får kyrkkontot:

1. Logga in på `https://m365.cloud.microsoft/`.
2. Kontrollera att **Excel** och **OneDrive** går att öppna från applistan.
3. Kontrollera separat om **SharePoint** finns. Utgå inte från att tjänsten ingår bara för att det är ett ideellt Microsoft 365-paket.
4. Öppna `https://admin.microsoft.com/`.
5. Gå till **Billing > Your products** och anteckna paketets exakta namn.
6. Öppna `https://entra.microsoft.com/`.

Det viktiga är inte paketnamnet. Följande funktioner måste finnas:

- Excel för webben
- OneDrive
- Microsoft Entra ID
- Möjlighet att skapa en appregistrering

Microsoft Entra ID Free följer med Microsoft 365. Appregistreringen och normal användning av Microsoft Graph har ingen separat avgift för detta användningsfall. En betald Azure-prenumeration behövs inte.

Om du inte kan öppna **App registrations** i Entra kan vanliga användares rätt att registrera appar vara avstängd. Då behövs en person med rollen **Application Administrator**, **Cloud Application Administrator** eller **Global Administrator** för engångssteget i avsnitt 4.

## Steg 2: lägg arbetsboken i kyrkans OneDrive

Eftersom OneDrive säkert finns börjar installationen där. Använd ett kyrkägt funktions- eller administrationskonto som kyrkan kan behålla när personer byts ut. Använd inte Daniels privata Microsoft-konto och dela inte kontots lösenord mellan användarna.

1. Logga in på Microsoft 365 med det kyrkkonto som ska äga arbetsboken.
2. Öppna OneDrive.
3. Skapa mappen **Inventarier**.
4. Ladda upp `workbook/Lutherska-Inventarier.xlsx` från projektet.

Om kontrollen i steg 1 visar att SharePoint ingår kan arbetsboken senare flyttas till en kyrkägd SharePoint-dokumentyta. Det ger bättre organisationsägarskap, men är inte ett krav för att börja.

## Steg 3: ge användarna åtkomst till arbetsboken

1. Markera `Lutherska-Inventarier.xlsx` i SharePoint eller OneDrive.
2. Välj **Manage access/Hantera åtkomst**.
3. Ge användarna eller en lämplig Microsoft 365-grupp behörigheten **Can edit/Kan redigera**.
4. Använd inte en anonym länk av typen **Anyone/Vem som helst**.
5. Välj **Copy link/Kopiera länk**.
6. Välj helst **People with existing access/Personer med befintlig åtkomst**. Länken ska inte i sig ge nya personer behörighet.
7. Spara länken. Den behövs som `M365_WORKBOOK_URL`.

Varje person måste använda en egen identitet som kyrkans Microsoft 365-tenant känner igen och ha redigeringsbehörighet till filen. Appen kan inte ge någon mer behörighet än användaren redan har.

### Vilka konton kan användarna ha?

Det finns två praktiska alternativ:

1. **Kyrkkonton i Microsoft 365** är det enklaste och det som ska användas i första testet. Kontona behöver kunna logga in och öppna den delade arbetsboken. Alla medlemmar behöver inte automatiskt få ett sådant konto.
2. **Privata Microsoft-konton som gäster** kan fungera om kontot först bjuds in som gästanvändare till kyrkans Microsoft Entra-tenant och arbetsboken delas till gästen. Den här vägen är mer administrativ och ska provas med exakt ett konto innan den används bredare.

Ett privat Outlook-, Hotmail- eller Gmail-konto kan alltså inte bara skriva in sina vanliga uppgifter i appen. Det måste först finnas som gäst i kyrkans tenant. Gmail-användare kan beroende på tenantens inställningar logga in via engångskod eller kopplat Microsoft-konto.

Microsofts ideella licenser är avsedda för behörig personal och vissa volontärscenarier; vanliga medlemmar eller deltagare ska inte automatiskt tilldelas ideella licenser. Gäståtkomst är därför den tänkbara vägen för en medlem som saknar kyrkkonto. För den första driftsättningen rekommenderas ändå två befintliga kyrkkonton, eftersom det minskar antalet osäkra delar.

## Steg 4: registrera appen i Microsoft Entra

1. Öppna `https://entra.microsoft.com/` med kyrkkontot.
2. Gå till **Identity > Applications > App registrations**.
3. Välj **New registration**.
4. Ange namnet **Lutherska Inventarier**.
5. Välj **Accounts in this organizational directory only**.
6. Under **Redirect URI** väljer du plattformen **Single-page application (SPA)**.
7. Ange exakt:

   `https://danielomazarino.github.io/lutherska-inventarier/`

8. Välj **Register**.
9. På sidan **Overview**, kopiera:
   - **Application (client) ID**
   - **Directory (tenant) ID**
10. Gå till **Authentication** och lägg även till följande SPA-adress för lokal testning:

    `http://localhost:5173/`

11. Gå till **API permissions > Add a permission > Microsoft Graph > Delegated permissions**.
12. Sök efter och lägg till **Files.ReadWrite**.
13. Skapa ingen klienthemlighet och lägg inte till application permissions.

`Files.ReadWrite` kräver enligt Microsoft normalt inte administratörsgodkännande. Kyrkans policy kan ändå blockera användargodkännande. Om inloggningen visar **Need admin approval** ska en administratör öppna appregistreringens **API permissions** och välja **Grant admin consent**.

### Säkerhetsinnebörden

- Appen får aldrig användarens lösenord.
- Inloggningen sker hos Microsoft via MSAL.
- Behörigheten är delegerad: appen agerar endast medan en riktig användare är inloggad.
- Koden använder endast den konfigurerade arbetsboken.
- Den tekniska behörigheten `Files.ReadWrite` är dock inte låst till just en fil. Den gäller användarens filer enligt Microsofts behörighetsmodell.
- Använd därför endast kyrkkonton som behöver registret och ge dem endast de SharePoint-rättigheter de behöver.

## Steg 5: konfigurera GitHub Pages utan kyrkkonto

Det här gör du på ditt privata GitHub-konto.

1. Öppna repositoryt `danielomazarino/lutherska-inventarier`.
2. Gå till **Settings > Secrets and variables > Actions > Variables**.
3. Skapa följande tre **Repository variables**:

| Namn | Värde |
| --- | --- |
| `M365_TENANT_ID` | Directory (tenant) ID från Entra |
| `M365_CLIENT_ID` | Application (client) ID från Entra |
| `M365_WORKBOOK_URL` | Delningslänken från SharePoint/OneDrive |

4. Öppna fliken **Actions**.
5. Välj **Deploy to GitHub Pages**.
6. Välj **Run workflow > Run workflow**.
7. Vänta tills både build och deploy är gröna.
8. Öppna `https://danielomazarino.github.io/lutherska-inventarier/`.

Tenant ID och Client ID är publika identifierare, inte lösenord. Arbetsbokslänken byggs också in i den publika appen och ska därför inte vara en anonym delningslänk. En utomstående kan inte läsa filen utan både Microsoft-inloggning och faktisk behörighet till arbetsboken.

## Steg 6: första riktiga inloggningen

1. Öppna webbadressen i ett privat webbläsarfönster.
2. Välj **Anslut Microsoft 365**.
3. Kontrollera att Tenant ID, Client ID och arbetsbokslänken redan är ifyllda.
4. Välj **Logga in och anslut**.
5. Logga in med kyrkkontot.
6. Godkänn `Files.ReadWrite` om Microsoft frågar.
7. Kontrollera att vänsterspalten visar **Synkad med Excel** och kontots namn.

Om appen fortfarande visar **Sampleläge**, kontrollera först GitHub Actions-variablerna och kör deployment-arbetsflödet igen.

## Steg 7: verifiera hela arbetsflödet

Gör testen i denna ordning och kontrollera arbetsboken efter varje steg:

1. Lägg till kategorin **Textilier** under **Administration**.
2. Lägg till platsen **Sakristian**.
3. Lägg till ett testföremål med märkningen **TEST-001**.
4. Öppna arbetsboken i Excel för webben och kontrollera att raden finns i tabellen `Inventory`.
5. Låna ut föremålet och ange både låntagare och vem som registrerar lånet.
6. Kontrollera att en ny rad finns i tabellen `Loans`.
7. Kontrollera att föremålet inte längre går att låna ut en gång till i appen.
8. Ändra testlånets `DueAt` i Excel till gårdagens datum.
9. Gå tillbaka till appen och välj **Uppdatera**.
10. Kontrollera att lånet visas som försenat i rött.
11. Markera föremålet som återlämnat.
12. Kontrollera att `ReturnedAt` fyllts i i Excel.
13. Genomför tvåanvändartestet nedan.

Radera därefter raderna för **TEST-001** från tabellerna `Loans` och `Inventory` i Excel. Ta endast bort dataraderna, inte tabeller, rubriker eller arbetsblad.

### Exakt tvåanvändartest

Testets syfte är att bevisa att informationen ligger i den gemensamma arbetsboken och inte endast i en persons webbläsare.

Förbered två separata identiteter:

- **Konto A**: kyrkkontot som äger arbetsboken
- **Konto B**: ett annat kyrkkonto med redigeringsbehörighet till arbetsboken

Använd två olika webbläsarprofiler, två olika datorer eller en vanlig och en privat webbläsare. Det räcker inte att öppna två flikar i samma profil, eftersom Microsoft då normalt återanvänder samma inloggning.

1. Logga in som Konto A och kontrollera att appen visar **Synkad med Excel** och namnet för Konto A.
2. Logga in separat som Konto B och kontrollera att appen visar Konto B, inte Konto A.
3. Lägg till föremålet **TEST-A** som Konto A.
4. Klicka **Uppdatera** som Konto B. Föremålet ska visas där.
5. Registrera ett lån av **TEST-A** som Konto B.
6. Klicka **Uppdatera** som Konto A. Föremålet ska visas som utlånat och lånet ska finnas i utlåningsloggen.
7. Markera lånet som återlämnat som Konto A.
8. Klicka **Uppdatera** som Konto B. Återlämningen ska visas där.
9. Öppna Excel-arbetsboken och kontrollera att samma rader finns i tabellerna `Inventory` och `Loans`.

Skrivningar från appen sparas direkt i Excel. Andra redan öppna webbläsare uppdateras automatiskt ungefär en gång per minut eller direkt när användaren väljer **Uppdatera**. Det är alltså gemensam data med regelbunden uppdatering, inte en sekundsnabb realtidskanal.

När testet med två kyrkkonton fungerar kan Konto B ersättas med ett privat konto som först bjudits in som gäst. Upprepa då hela testet. Betrakta inte privata medlemskonton som stödda förrän just den gästinloggningen har fungerat mot kyrkans tenant och arbetsbok.

## Varför detta är bättre än ännu ett kalkylblad

Det är fortfarande Excel som lagrar informationen. Skillnaden är att Excel inte längre är det dagliga användargränssnittet.

| Vanligt kalkylblad | Inventarieappen |
| --- | --- |
| Användaren måste förstå flikar, kolumner och relationer | Användaren ser arbetsuppgifter: lägg till, låna ut, lämna tillbaka |
| Kategori, grupp och plats skrivs ofta olika | Val görs från gemensamma listor |
| Det är svårt att se aktuell status bland många rader | Översikten visar utlånat och försenat direkt |
| Låntagare och registrerare blandas lätt ihop | Båda uppgifterna är separata och obligatoriska |
| Samma föremål kan råka registreras som utlånat flera gånger | Utlåningsknappen stängs av när föremålet redan är utlånat |
| Formler, filter eller rubriker kan ändras av misstag | Vanliga användare behöver inte röra tabellstrukturen |
| Mobil användning kräver panorering i breda tabeller | Appen är anpassad för iPad och dator |
| Varje ny grupp skapar gärna en ny fil eller egen struktur | Alla arbetar mot samma register och samma klassificeringar |

Det viktiga värdet är alltså inte att appen "ersätter Excel med något dyrt". Den lägger ett kontrollerat arbetsflöde ovanpå ett verktyg kyrkan redan har och betalar för.

## Vad lösningen inte löser ännu

Var öppen med följande under drift:

- Excel är fortfarande en relativt känslig datakälla vid många samtidiga skrivningar.
- Personer med redigeringsbehörighet kan fortfarande öppna filen och ändra strukturen manuellt.
- Appen skickar ännu inga e-postpåminnelser. Försenat syns när appen öppnas.
- Föremål kan läggas till men ännu inte redigeras eller tas bort i appen.
- GitHub Pages-webbplatsen drivs från Daniels privata konto. Kyrkan äger arbetsboken och appregistreringen, men inte webbadressen.

Detta är ett rimligt kostnadsfritt upplägg för en mindre verksamhet. Om användningen växer är nästa rimliga steg att flytta datan till Microsoft Lists, som ofta ingår i samma Microsoft 365-paket. Appens arbetsflöde kan behållas samtidigt som datalagringen blir mindre känslig.

## Kostnad och långsiktigt ägarskap

Lösningens extra kostnad är 0 kronor om kyrkans befintliga paket innehåller Excel, SharePoint/OneDrive och Entra ID:

- GitHub Pages för ett publikt repository: kostnadsfritt
- Microsoft Graph för normal användning: ingen separat avgift
- Excel-arbetsboken: ingår i befintligt Microsoft 365-paket
- Separat server eller databas: ingen

Om kyrkan senare vill äga även webbadressen kan webbappen flyttas till kyrkans befintliga webbhotell under exempelvis `https://lutherska.nu/inventarier/`. Det kräver normalt inte ett nytt webbhotell, men någon med åtkomst till webbplatsens filer måste publicera innehållet i mappen `dist`. När adressen ändras måste den nya adressen också läggas till som SPA redirect URI i Entra.

Spara följande i kyrkans egen administrativa dokumentation:

- Appregistreringens namn, Tenant ID och Client ID
- Var arbetsboken ligger
- Vilka som har åtkomst
- Webbadressen
- Vem som ansvarar för GitHub-publiceringen
- Den här installationsguiden

## Tekniskt underhåll av appkoden

Det här avsnittet behövs **inte** för kyrkans vanliga användning. När appen väl ligger på GitHub Pages behöver ordföranden och administratörerna inte installera Node.js, köra kommandon eller starta appen lokalt. De öppnar bara webbadressen.

Kommandona nedan behövs endast om någon senare ändrar appens källkod eller bygger om Excel-mallen. GitHub Pages bygger och publicerar automatiskt en ny version när en ändring skickas till GitHub-projektets huvudgren.

För lokalt utvecklingsarbete behöver den tekniskt ansvariga personen Node.js och npm. Första gången körs:

```powershell
npm install
npm run dev
```

`npm install` hämtar de programdelar som appen använder. `npm run dev` öppnar en lokal testversion på datorn. Den lokala testversionen påverkar inte den publicerade appen.

Före publicering bör följande kontroller köras:

```powershell
npm run workbook:verify-groups
npm run lint
npm run build
```

- `workbook:verify-groups` kontrollerar att verksamhetsgrupperna är samma i appen och Excel-mallen.
- `lint` söker efter vanliga kodfel.
- `build` kontrollerar att den färdiga webbappen kan skapas.

Två ytterligare kommandon används bara vid särskilda ändringar:

- `npm run workbook:create` skapar om den tomma Excel-mallen.
- `npm run workbook:render-preview` skapar underlaget för arbetsboksbilderna i dokumentationen.

Ingen av dessa åtgärder krävs för att registrera inventarier, hantera lån eller administrera användarnas åtkomst till Excel-filen.

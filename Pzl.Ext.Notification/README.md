# Pzl.Ext.Notification – Teams-varsling for driftsmeldinger og nyheter

Valgfritt tillegg til [Driftsmeldinger (Pzl.Part.Announcement)](../Pzl.Part.Announcement/README.md). En timer-trigget Azure Function sender varsler i Teams-aktivitetsfeeden (bjella) når

- en **nyhetsside** publiseres på et konfigurert område, og
- en **driftsmelding** blir aktiv og enten har et meldingsnivå som er satt opp under «Nivåer som alltid varsles» for området, eller har «Varsle i Teams» satt. Driftsmeldinger som allerede er *Løst*, varsles ikke, og
- en redaktør setter **«Varsle på nytt»** på en driftsmelding (ved vesentlige endringer, eller når saken er løst).

Tittelen i bjella sier hva slags varsel det er:

| Varsel | Tittel i bjella |
|---|---|
| Første varsel, status *Pågår* | «Driftsmelding: <tittel>» |
| Første varsel, status *Planlagt* | «Planlagt: <tittel>» |
| «Varsle på nytt», status *Løst* | «Løst: <tittel>» |
| «Varsle på nytt», ellers | «Oppdatert: <tittel>» |
| Nyhet | «<tittel>» |

Varslene respekterer **målgruppestyring**: har elementet målgrupper, varsles medlemmene i de gruppene. Ellers varsles standardgruppene for området. Klikk på et driftsmeldingsvarsel åpner en Teams-fane som viser intranettet, med meldingen i popover. Klikk på et nyhetsvarsel åpner samme fane, og customizer-en **NotificationDeepLink** i SPFx-pakken til dette tillegget (se [steg 4](#steg-4-teams-appen)) sender brukeren rett til nyhetssiden. Stage view ble prøvd for nyheter og forkastet: artikkelen åpnet i et vindu uten intranettets navigasjon, og mobil støtter det ikke. Graph krever at lenkene er Teams-deep-linker, så fanen er enten varsel-appens egen fane (standard) eller en fane i intranettets egen Teams-app.

Lista og webdelen fungerer uten dette tillegget. Kolonnene tjenesten bruker (`PzlNotifyTeams` «Varsle i Teams», `PzlNotifyAgain` «Varsle på nytt» og `PzlNotifiedAt` «Varslet i Teams») legges til driftsmeldingslista med malen `Templates/AnnouncementFields.xml` i dette tillegget, se [steg 3](#steg-3-sharepoint).

## Innhold

| Mappe | Hva |
|---|---|
| `src/`, `config/`, `sharepoint/` | SPFx-pakken `pzl-ext-notification.sppkg` med customizer-en **NotificationDeepLink** (videresending av nyhetsvarsler). |
| `function/` | Azure Function (TypeScript, Node 20+, Functions v4). Kjører hvert minutt. |
| `teams-app/` | Manifest-mal og ikoner for Teams-appen som eier varslene. |
| `scripts/` | `Install-Notification.ps1` (alt under i ett, se [Installer med skript](#installer-med-skript)), `Deploy-NotificationInfrastructure.ps1` (Azure-ressurser), `Grant-NotificationPermissions.ps1` (app-rettigheter), `New-NotificationTeamsApp.ps1` (Teams-app-pakke), `Publish-NotificationFunction.ps1` (bygg og deploy) og `Get-NotificationLog.ps1` (logg fra Application Insights). |
| `Templates/AnnouncementFields.xml` | PnP-mal som legger varslingskolonnene til driftsmeldingslista. |
| `Templates/NotificationConfig.xml` | PnP-mal for konfigurasjonslista `Varslingskonfigurasjon`. |
| `Templates/DeepLinkCustomAction.xml` | PnP-mal som registrerer **NotificationDeepLink** på ett område (ved tenant-distribusjon). |

## Slik virker det

```
Timer (1 min)
  └─ Les aktive rader i Varslingskonfigurasjon (CONFIG_SITE_URL)
       └─ Per område
            ├─ Driftsmeldinger: PzlStartDate <= nå, PzlEndDate >= nå, og enten
            │     • PzlNotifyAgain = ja (nytt varsel), eller
            │     • PzlNotifiedAt tom, PzlStatus ikke Løst, og (PzlSeverity i radens
            │       «Nivåer som alltid varsles» eller PzlNotifyTeams = ja) (første varsel)
            │     → sett PzlNotifiedAt og PzlNotifyAgain = nei (claim) → send
            │     → ved feil: sett feltene tilbake
            └─ Nyheter: PromotedState = 2, FirstPublishedDate innen NEWS_LOOKBACK_DAYS,
                  publisert (hovedversjon, godkjent), ikke registrert i Table Storage
                  → registrer i tabellen (claim) → send → ved feil: fjern raden
```

- Mottakere: `transitiveMembers` på gruppene via Graph, dedupliserte, uten deaktiverte kontoer og (som standard) uten gjester.
- Sending: `POST /teamwork/sendActivityNotificationToRecipients` i grupper på 100, med retry ved 429/5xx.
- Nyhetssider skrives aldri til. En feltoppdatering på en publisert side ville gitt en ny versjon og «upubliserte endringer», så sendt-status for sider ligger i Table Storage i Function-appens lagringskonto.
- Timer-triggere kjører som singleton, og claim-før-send gjør at overlappende kjøringer uansett ikke dobbeltsender.

## Forutsetninger

- Driftsmeldinger (Pzl.Part.Announcement) 2.0.0 eller nyere, med hovedmalen kjørt. Med eldre lister uten *Status* behandles alle driftsmeldinger som *Pågår*.
- Azure-abonnement, og rettigheter til å opprette ressurser og tildele app-roller i Entra (Global Administrator eller Privileged Role Administrator for steg 2).
- Verktøy: [Azure CLI](https://learn.microsoft.com/cli/azure/), [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) v4, Node 20 eller 22, PnP.PowerShell, Microsoft.Graph PowerShell.

## Installer med skript

`scripts/Install-Notification.ps1` gjør steg 1, 2, 4 og 5 under, pluss tilgangen per område fra steg 2, i én kjøring. Skriptet kan kjøres på nytt ved hver oppgradering og installerer ingen verktøy. SharePoint-delen (steg 3 og SPFx-pakken) gjøres med `Pzl.Part.Announcement/scripts/Install-Driftsmeldinger.ps1 -EnableTeamsNotifications`.

```powershell
az login --tenant <tenant-id>
.\scripts\Install-Notification.ps1 -TenantId <tenant-id> -SubscriptionId <abonnement> `
  -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification `
  -StorageAccountName stintranetnotify01 -ConfigSiteUrl https://<tenant>.sharepoint.com `
  -AnnouncementSiteUrls https://<tenant>.sharepoint.com -NewsSiteUrls https://<tenant>.sharepoint.com/sites/nyheter `
  -AppName "<Intranett> varsler" -DeveloperName "<organisasjon>"
```

| Steg | Hva skriptet gjør |
|---|---|
| 1. Azure | Kjører `Deploy-NotificationInfrastructure.ps1`. Teams-app-ID-en, `DRY_RUN` og listetitlene leses fra den eksisterende Function-appen og beholdes. `-DryRun On/Off` overstyrer. Nye installasjoner starter med `DRY_RUN=true`. |
| 2. Entra | Kjører `Grant-NotificationPermissions.ps1` med identitetens principal ID. |
| Tilgang per område | `Write` på `-AnnouncementSiteUrls` (standard: konfigurasjonsområdet), `Read` på konfigurasjonsområdet og `-NewsSiteUrls`. Eksisterende tilganger heves, men senkes aldri. Bruker PnP.PowerShell (lastet i økten, via `Import-PnPPowerShell` eller installert) med PP365-appen som standard `-ClientId`. |
| 4. Teams-appen | Bygger pakken med `New-NotificationTeamsApp.ps1` og laster den opp til organisasjonens appkatalog med Graph: ny app første gang, ny versjon (patch-nummeret økes) ved oppgradering. `-SkipTeamsAppUpload` bygger bare pakken. |
| 5. Funksjonen | Kjører `Publish-NotificationFunction.ps1`. |

Roller: Owner (eller Contributor + User Access Administrator) på abonnementet, Global Administrator eller Privileged Role Administrator, Teams Administrator og SharePoint Administrator. Hvert steg kan hoppes over med `-Skip…`. For hånd gjenstår app setup policy i Teams admin center, radene i **Varslingskonfigurasjon** og testvarselet. Skriptet skriver dem ut til slutt.

## Steg 1: Azure-ressurser

`scripts/Deploy-NotificationInfrastructure.ps1` oppretter alt som trengs, idempotent (kan kjøres på nytt): ressursgruppe, user-assigned managed identity, lagringskonto, Application Insights, rolletildelinger for identiteten på lagringskontoen, Function-app på Flex Consumption (Node 22) med identitetsbasert tilgang til lagring, og app-innstillingene fra steg 5. Krever Azure CLI 2.70+ innlogget med rettigheter til å opprette ressurser og rolletildelinger.

```powershell
az login --tenant <tenant-id>
.\scripts\Deploy-NotificationInfrastructure.ps1 `
  -SubscriptionId <abonnement> `
  -ResourceGroupName rg-intranet-notification `
  -Location norwayeast `
  -FunctionAppName func-intranet-notification `
  -StorageAccountName stintranetnotify01 `
  -ConfigSiteUrl https://<tenant>.sharepoint.com
```

Standardregionen er Norway East. Regionen må støtte Flex Consumption (`az functionapp list-flexconsumption-locations`), og skriptet sjekker det før noe opprettes. Azure kan også avvise en region for abonnementet («The selected region is currently not accepting new customers», sett i West Europe). Velg da en annen med `-Location`, f.eks. `swedencentral`. Valgfrie parametere: `-TeamsDeepLinkAppId`, `-TeamsDeepLinkEntityId`, `-ConfigListTitle`, `-AnnouncementListTitle`, `-DryRun` (standard `$true`).

Scriptet skriver ut **client ID** og **principal ID** for identiteten til slutt. De brukes i de neste stegene.

## Steg 2: Rettigheter i Entra

Tjenesteidentiteten trenger disse applikasjonstillatelsene:

| Ressurs | Tillatelse | Brukes til |
|---|---|---|
| Microsoft Graph | `TeamsActivity.Send` | Sende aktivitetsvarsler |
| Microsoft Graph | `GroupMember.Read.All` | Finne medlemmer i mottakergruppene |
| Microsoft Graph | `Sites.Selected` | Kreves sammen med SharePoint-rollen under |
| SharePoint Online | `Sites.Selected` | Lese lister via SharePoint REST, skrive `PzlNotifiedAt` |

```powershell
.\scripts\Grant-NotificationPermissions.ps1 -TenantId <tenant-id> -ServicePrincipalObjectId <principal-id fra steg 1>
```

Deretter gis tilgang per område. Konfigurasjonsområdet trenger `Read`, hvert område med driftsmeldinger trenger `Write` (for `PzlNotifiedAt`), og områder med bare nyheter trenger `Read`:

```powershell
Connect-PnPOnline -Url https://<tenant>.sharepoint.com -Interactive -ClientId <pnp-client-id>
Grant-PnPAzureADAppSitePermission -AppId <MI_CLIENT> -DisplayName "Intranett-varsling" -Site https://<tenant>.sharepoint.com -Permissions Write
Grant-PnPAzureADAppSitePermission -AppId <MI_CLIENT> -DisplayName "Intranett-varsling" -Site https://<tenant>.sharepoint.com/sites/<annet-omraade> -Permissions Read
```

## Steg 3: SharePoint

Steg 1 til 3 under (og SPFx-pakken fra steg 4) kan gjøres med `Pzl.Part.Announcement/scripts/Install-Driftsmeldinger.ps1 -EnableTeamsNotifications`, se [README for Driftsmeldinger](../Pzl.Part.Announcement/README.md#with-the-install-script-new-installation-and-upgrade). Radene i konfigurasjonslista og målgruppestyring gjøres for hånd.

1. Legg varslingskolonnene til driftsmeldingslista på hvert område tjenesten skal varsle fra. Malen bygger på innholdstypen fra hovedmalen, så kjør den etter `Announcement.xml`, og kjør den på nytt hver gang hovedmalen er kjørt på nytt:

   ```powershell
   Connect-PnPOnline -Url https://<tenant>.sharepoint.com -Interactive -ClientId <pnp-client-id>
   Invoke-PnPSiteTemplate -Path .\Templates\AnnouncementFields.xml
   ```

   Slå deretter på «Varsle i Teams» under «Felter i skjema for ny driftsmelding» i egenskapene til webdelen, hvis redaktørene registrerer driftsmeldinger der. Feltet er skjult i webdelens skjema som standard. «Varsle på nytt» vises bare i listeskjemaet når et element redigeres. Legg gjerne kolonnene til visningen av lista.
2. Opprett konfigurasjonslista på området som skal eie konfigurasjonen (typisk rot- eller hub-området):

   ```powershell
   Connect-PnPOnline -Url https://<tenant>.sharepoint.com -Interactive -ClientId <pnp-client-id>
   Invoke-PnPSiteTemplate -Path .\Templates\NotificationConfig.xml
   ```

3. Legg inn én rad per område i **Varslingskonfigurasjon**:

   | Kolonne | Verdi |
   |---|---|
   | Tittel | Visningsnavn, brukes i loggen (og under varselet hvis områdets tittel ikke kan leses) |
   | Områdeadresse | Full URL til området |
   | Varsle nyheter | Ja/nei |
   | Varsle driftsmeldinger | Ja/nei |
   | Nivåer som alltid varsles | Meldingsnivåer som varsles uten at «Varsle i Teams» er satt på driftsmeldingen, typisk «Kritisk». Tom betyr at driftsmeldinger bare varsles når «Varsle i Teams» er satt. Verdiene tolkes uavhengig av språk, og 1.x-verdien «Varsel» betyr «Kritisk». |
   | Standard mottakergrupper | Objekt-ID for én eller flere Entra-grupper, én per linje. Brukes når elementet ikke har målgruppe. |
   | Driftsmeldingsliste | Tom for standard (`Driftsmeldinger`), ellers tittel på lista |
   | Aktiv | Ja |

   Et område med egen mottakergruppe (for eksempel en lederportal) får en egen rad med den gruppa som standard.

   Mangler kolonnen «Nivåer som alltid varsles», eller er den tom, varsles driftsmeldinger bare når «Varsle i Teams» er satt.

4. Aktiver målgruppestyring på Site Pages-biblioteket og driftsmeldingslista der målgrupper skal brukes. Uten målgruppestyring brukes alltid standardgruppene.

## Steg 4: Teams-appen

Aktivitetsvarsler må sendes fra en Teams-app der `webApplicationInfo.id` er client ID til avsenderen. Intranett-appen som bruker SharePoint-SSO kan ikke endre det feltet, derfor en egen, minimal app.

1. Bygg pakken med `scripts/New-NotificationTeamsApp.ps1`. Den fyller inn malen i `teams-app/manifest.json` og zipper den med ikonene til `teams-app/dist/`:

   ```powershell
   .\scripts\New-NotificationTeamsApp.ps1 `
     -NotificationClientId <client-id fra steg 1> `
     -SiteUrl https://<tenant>.sharepoint.com `
     -AppName "Intranett-varsler" `
     -DeveloperName "<organisasjon>"
   ```

   Anbefalte verdier i produksjon:

   | Parameter | Anbefaling |
   |---|---|
   | `-AppName` | Intranettets navn pluss «varsler», f.eks. «Svaneposten varsler». Navnet vises i bjella og i applisten. |
   | `-DeveloperName` | Organisasjonen som eier intranettet, f.eks. kommunens navn. |
   | `-ColorIcon` / `-OutlineIcon` | Intranett-appens egne ikoner (192x192 farge, 32x32 transparent), så varslene får samme ikon som intranettet. |
   | `-AppId` | Ny GUID første gang. Ta vare på den og bruk samme ved oppdateringer. |
   | `-Version` | Øk ved hver opplasting, f.eks. 1.0.0, 1.0.1. |

   Tekstene i `activityTypes` i `teams-app/manifest.json` er det brukeren ser i bjella og kan tilpasses før bygging. Begge typene viser bare `{title}`: for driftsmeldinger setter tjenesten selv inn prefikset («Driftsmelding:», «Planlagt:», «Oppdatert:» eller «Løst:», se tabellen øverst), og for nyheter framgår typen av aktivitetstypens beskrivelse («Nyhet») og forhåndsvisningen. Linjen under varselet viser tittelen på området varselet kommer fra (områdets navn i SharePoint, lest fra `web/Title` ved hver kjøring). Kan ikke tittelen leses, brukes tittelen på konfigurasjonsraden i steg 3 i stedet.

2. Last opp zip-fila i Teams admin center (Teams apps > Manage apps > Upload new app).
3. Lag en **app setup policy** som installerer appen for mottakerne (uten å feste den i appfeltet), og tildel den til samme grupper som i konfigurasjonen. Start med en testgruppe.
4. Pek varslene på en Teams-fane. Graph godtar bare Teams-deep-linker som `webUrl`, så dette er påkrevd. Enklest er varsel-appens egen fane, som åpner intranettet:

   ```bash
   az functionapp config appsettings set -g <ressursgruppe> -n <function-app> --settings TEAMS_DEEPLINK_APP_ID=<Teams-app-ID fra punkt 1> TEAMS_DEEPLINK_ENTITY_ID=intranet
   ```

   Fanen i varsel-appen sender `{subEntityId}` videre til intranettet som `?driftsmelding=`, og det er den parameteren webdelen bruker: `driftsmelding:<id>` åpner driftsmeldingen i popover, `page:<url>` (nyheter) sender brukeren videre til nyhetssiden. Webdelen må derfor ligge på siden fanen åpner, normalt forsiden.

   Videresendingen for nyheter gjøres av application customizer-en **NotificationDeepLink** (komponent-ID `1e8a8edf-84b4-4a5e-9fcc-5ec33e34a509`) i SPFx-pakken til dette tillegget, så den skjer før sideinnholdet lastes og uavhengig av hvilke webdeler siden har. Bygg pakken med `npm i` og `npm run package` i denne mappa (Node 22), og last opp `sharepoint/solution/pzl-ext-notification.sppkg` til appkatalogen. Customizer-en registreres som custom action på området når pakken legges til fra en områdesamlings-appkatalog. Er pakken distribuert for hele tenanten, kjør malen på området fanen åpner:

   ```powershell
   Connect-PnPOnline -Url https://<tenant>.sharepoint.com -Interactive -ClientId <pnp-client-id>
   Invoke-PnPSiteTemplate -Path .\Templates\DeepLinkCustomAction.xml
   ```

   Sjekk med `Get-PnPCustomAction -Scope Site`. Customizer-en gjør ingenting på sider som åpnes uten `driftsmelding=page:`-parameteren. Bare samme-opprinnelse `.aspx`-stier godtas, og andre URL-parametere (som `app=portals` fra Teams-fanen) beholdes.

   **Har intranettet en egen Teams-app** med SharePoint-SSO, pek varslene dit i stedet. Det er anbefalt der den finnes, siden varsel-appen ikke garanterer stille pålogging på desktop og mobil:

   1. Legg `{subEntityId}` inn i fanens `contentUrl` i intranett-appens manifest, som URL-parameter `driftsmelding` på siden `dest` peker til. Har siden allerede parametere, legg den til med `%26` (URL-kodet `&`):

      ```text
      https://<tenant>.sharepoint.com/_layouts/15/teamslogon.aspx?spfx=true&dest=https%3A%2F%2F<tenant>.sharepoint.com%2F%3Fapp%3Dportals%26driftsmelding%3D{subEntityId}
      ```

   2. Øk versjonen, last opp manifestet som oppdatering, og la brukerne logge ut og inn av Teams under test (se punkt 5).
   3. Sett `TEAMS_DEEPLINK_APP_ID` til intranett-appens app-ID og `TEAMS_DEEPLINK_ENTITY_ID` til fanens `entityId`.

   Varsel-appen må uansett finnes og være installert for mottakerne: Graph krever at avsenderens identitet står i `webApplicationInfo.id` i appen varselet vises for, og intranett-appen må ha SharePoints app-ID der for SSO. `subPageId` fra Teams-konteksten når ikke fram til webdelen i fanen, så URL-parameteren er eneste pålitelige vei.

5. Ved oppdatering av manifestet (ny versjon lastet opp i admin center): Teams-klientene beholder den gamle app-definisjonen til de henter på nytt, som kan ta timer. Logg ut og inn av Teams for å tvinge det fram under test.

> **Ikke verifisert i alle tenanter:** at en managed identity aksepteres som `webApplicationInfo.id`. Test tidlig med én mottaker. Alternativet er en app-registrering med managed identity som federated identity credential; Function-koden er den samme (`AZURE_CLIENT_ID` settes til app-registreringens client ID).

## Steg 5: Deploy Function-appen

`scripts/Publish-NotificationFunction.ps1` bygger prosjektet, pakker det og deployer med Azure CLI zip-deploy og remote build. Azure Functions Core Tools trengs ikke.

```powershell
.\scripts\Publish-NotificationFunction.ps1 -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification
```

App-innstillingene ble satt i steg 1, bortsett fra `TEAMS_DEEPLINK_APP_ID`, som settes i steg 4 når Teams-appen finnes. Endre innstillinger senere med `az functionapp config appsettings set`.

| Innstilling | Standard | Beskrivelse |
|---|---|---|
| `AZURE_CLIENT_ID` | | Client ID til managed identity (settes i steg 1) |
| `CONFIG_SITE_URL` | *(påkrevd)* | Område med konfigurasjonslista |
| `CONFIG_LIST_TITLE` | `Varslingskonfigurasjon` | Tittel på konfigurasjonslista |
| `ANNOUNCEMENT_LIST_TITLE` | `Driftsmeldinger` | Standard tittel på driftsmeldingslista |
| `NEWS_LOOKBACK_DAYS` | `7` | Hvor langt tilbake nye nyheter hentes |
| `INCLUDE_GUESTS` | `false` | Ta med gjestebrukere som mottakere |
| `TEAMS_DEEPLINK_APP_ID` | *(påkrevd)* | Teams-app som varslene åpner: varsel-appen selv (steg 4) eller intranett-appen. Uten denne sendes ingenting, og loggen sier hvorfor. |
| `TEAMS_DEEPLINK_ENTITY_ID` | `intranet` | `entityId` for fanen i den appen. `intranet` er varsel-appens fane. |
| `TEAMS_NOTIFICATION_APP_ID` | *(tom)* | Sendes til Graph som `teamsAppId`. Trengs bare når en mottaker har flere Teams-apper med samme avsender-identitet (`409` «Found multiple applications with the same AAD App ID»). Prøv verdien med `&teamsAppId=` på test-endepunktet før den settes. |
| `STATE_STORAGE_ACCOUNT_NAME` | Function-appens konto | Lagringskonto for sendt-status for sider |
| `STATE_TABLE_NAME` | `NotificationState` | Tabellnavn |
| `TIMER_SCHEDULE` | `0 * * * * *` | NCRONTAB, hvert minutt |
| `DRY_RUN` | `false` | Logg hva som ville blitt sendt, uten å sende eller skrive |

Start med `DRY_RUN=true` og se i loggen at riktige elementer og mottakerantall plukkes opp. Flex Consumption har ingen logstream (`az webapp log tail` gir 404), så bruk `scripts/Get-NotificationLog.ps1`, som spør Application Insights, eller Invocations på funksjonen i portalen. Loggene kommer med ett til noen minutters forsinkelse.

```powershell
.\scripts\Get-NotificationLog.ps1 -ResourceGroupName rg-intranet-notification -FunctionAppName func-intranet-notification -Minutes 30
```

Sett så `DRY_RUN=false`.

## Steg 6: Test og pilot

Send først et testvarsel til én bruker med test-endepunktet i Function-appen. Det bruker Graphs per-bruker-endepunkt, som rapporterer leveringsfeil, i motsetning til masse-endepunktet timeren bruker:

```powershell
$key = az functionapp keys list -g rg-intranet-notification -n func-intranet-notification --query functionKeys.default -o tsv
$userId = az ad user show --id bruker@tenant.onmicrosoft.com --query id -o tsv
Invoke-RestMethod -Method Post -Uri "https://func-intranet-notification.azurewebsites.net/api/test-notification?userId=$userId&type=announcement" -Headers @{ 'x-functions-key' = $key } | ConvertTo-Json -Depth 6
```

Svaret inneholder `ok` og eventuell Graph-feil (`graphStatus` og `graphError`). `ConvertTo-Json` trengs for å se feilteksten, ellers viser PowerShell bare `@{error=}`. `403` betyr som regel at Teams-appen ikke er installert for brukeren ennå: vent på setup policy, eller legg til appen selv i Teams (Apper > Bygd for organisasjonen). Varselet skal dukke opp i bjella innen sekunder. `type=news` tester nyhetsvarianten. Legg til `&bulk=1` for å sende via masse-endepunktet, og `&teamsAppId=<id>` for å sende en Teams-app-ID til Graph (se `TEAMS_NOTIFICATION_APP_ID`). Et mislykket kall viser også hvilken app-ID, tenant og hvilke roller Graph-tokenet har (`token`). `&installationId=<id>` bruker brukerens installasjon av appen som emne i stedet for tekst og lenke (ID-en fra `GET /users/{id}/teamwork/installedApps`), for å skille feil i installasjonen fra feil i lenken.

> **Malparametere må stemme nøyaktig** med plassholderne i `templateText` i manifestet (`{title}`). Masse-endepunktet godtar ekstra parametere, men leverer da ingenting, uten feil. Per-bruker-endepunktet svarer «Incorrect template parameter arity». Endrer du `templateText`, må `templateParameters` i `announcements.ts`/`news.ts` endres tilsvarende.


1. Konfigurasjonsrad mot en testgruppe, setup policy mot samme gruppe.
2. Test variantene: nyhet, driftsmelding med et nivå fra «Nivåer som alltid varsles», driftsmelding med «Varsle i Teams», driftsmelding med status *Planlagt*, og «Varsle på nytt» på en driftsmelding som er satt til *Løst*.
3. Klikk varselet på desktop, web og mobil og se at riktig side/driftsmelding åpnes.
4. Sjekk at `PzlNotifiedAt` settes på driftsmeldingen og vises i webdelens popover.
5. Bytt standardgruppe og setup policy til produksjonsgruppene.

## Lokal kjøring

Kopier `function/local.settings.json.example` til `local.settings.json`. Lokalt brukes `DefaultAzureCredential`, så enten `az login` med en bruker som har tilgangene (ikke anbefalt for app-roller) eller en app-registrering med `AZURE_TENANT_ID`, `AZURE_CLIENT_ID` og `AZURE_CLIENT_SECRET` som har fått de samme rollene som i steg 2. Sett `STATE_STORAGE_ACCOUNT_NAME` til en ekte lagringskonto, Azurite støtter ikke identitetsbasert tilgang. Kjør `npm start`.

## Kjente funn fra `npm audit`

`@azure/identity` (siste 4.x) avhenger av `@azure/msal-node` 6.x, som har et høyt funn (GHSA-xcx6-jx8p-x3h5, rettet i msal-node 7.1.0). Sårbarheten ligger i authorization code-flyten for interaktiv pålogging, som tjenesten ikke bruker: i Azure autentiserer den med managed identity, lokalt med client credentials. Ingen utgave av `@azure/identity` bruker msal-node 7 ennå, og å tvinge det inn via `overrides` er en ustøttet kombinasjon. Kjør `npm audit` før deploy og oppdater `@azure/identity` når en versjon med msal-node 7 finnes.

## Feilsøking

| Symptom | Sjekk |
|---|---|
| `RequestDisallowedByAzure` («not accepting new customers») i steg 1 | Azure tar ikke imot nye kunder i regionen for dette abonnementet. Kjør på nytt med en annen `-Location`. Ressursgruppa som alt er opprettet, kan beholdes |
| `401`/`403` fra SharePoint | App-rollen `Sites.Selected` på **SharePoint Online** (ikke bare Graph), og `Grant-PnPAzureADAppSitePermission` på riktig område med riktig nivå |
| `403` fra `transitiveMembers` | `GroupMember.Read.All` tildelt, og at admin consent er gitt (scriptet tildeler app-roller direkte, consent er implisitt) |
| `400` «Invalid 'webUrl' specified» fra `sendActivityNotificationToRecipients` | `TEAMS_DEEPLINK_APP_ID` mangler eller er ikke en Teams-app-ID. Lenken må være `https://teams.microsoft.com/l/...` |
| `400`/`403` ellers fra `sendActivityNotificationToRecipients` | `TeamsActivity.Send`, at Teams-appen er installert for mottakeren (setup policy), og at `webApplicationInfo.id` i manifestet er client ID til identiteten som sender |
| `409` «Found multiple applications with the same AAD App ID» | Mottakeren har flere Teams-apper med avsender-identiteten i `webApplicationInfo.id`, typisk en eldre eller personlig opplastet utgave av varsel-appen. Fjern den gamle appen hos mottakeren (Teams > Apper > Administrer appene dine), eller sett `TEAMS_NOTIFICATION_APP_ID` |
| Ingen mottakere | Gruppe-ID-ene i konfigurasjonen, og at gruppa har medlemmer som ikke er deaktivert/gjester |
| Loggen sier «Sent … to N recipients», men ingenting i bjella | Appen var ikke installert for mottakeren da det ble sendt (setup policy tar tid), eller `templateParameters` matcher ikke `templateText`. Send et testvarsel med `/api/test-notification` til én bruker for å få feilen. |
| Nyheter varsles ikke | `PromotedState = 2`, publisert hovedversjon, innenfor `NEWS_LOOKBACK_DAYS`, og ikke allerede en rad i tabellen `NotificationState` |
| Driftsmeldinger varsles ikke | Startdato passert, sluttdato ikke passert, `PzlNotifiedAt` tom, status ikke *Løst*, og nivå/«Varsle i Teams». For nytt varsel: «Varsle på nytt» satt (kolonnen fra `AnnouncementFields.xml`) |
| Dobbelt varsel | Viva Connections kan også varsle om nyheter. Skru av det ene. |

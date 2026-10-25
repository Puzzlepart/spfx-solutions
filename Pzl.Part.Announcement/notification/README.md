# Teams-varsling for nyheter og driftsmeldinger

Valgfritt tillegg til [Pzl.Part.Announcement](../README.md). En timer-trigget Azure Function sender varsler i Teams-aktivitetsfeeden (bjella) når

- en **nyhetsside** publiseres på et konfigurert område, og
- en **driftsmelding** blir aktiv og enten har meldingsnivået «Varsel» eller har «Varsle i Teams» satt.

Varslene respekterer **målgruppestyring**: har elementet målgrupper, varsles medlemmene i de gruppene. Ellers varsles standardgruppene for området. Klikk på et varsel åpner en Teams-fane som viser intranettet, med elementet i fokus (driftsmeldinger åpnes i popover av webdelen). Graph krever at lenken er en Teams-deep-link, så fanen er enten varsel-appens egen fane (standard) eller en fane i intranettets egen Teams-app.

Lista og webdelen fungerer uten dette tillegget. Kolonnene `PzlNotifyTeams`, `PzlPinned` og `PzlNotifiedAt` finnes i hovedmalen fra versjon 1.1.0 og er uten effekt til tjenesten tas i bruk.

## Innhold

| Mappe | Hva |
|---|---|
| `function/` | Azure Function (TypeScript, Node 20+, Functions v4). Kjører hvert minutt. |
| `teams-app/` | Manifest-mal og ikoner for Teams-appen som eier varslene. |
| `scripts/` | `Deploy-NotificationInfrastructure.ps1` (Azure-ressurser), `Grant-NotificationPermissions.ps1` (app-rettigheter), `New-NotificationTeamsApp.ps1` (Teams-app-pakke), `Publish-NotificationFunction.ps1` (bygg og deploy) og `Get-NotificationLog.ps1` (logg fra Application Insights). |
| `../Templates/NotificationConfig.xml` | PnP-mal for konfigurasjonslista `Varslingskonfigurasjon`. |

## Slik virker det

```
Timer (1 min)
  └─ Les aktive rader i Varslingskonfigurasjon (CONFIG_SITE_URL)
       └─ Per område
            ├─ Driftsmeldinger: PzlStartDate <= nå, PzlEndDate >= nå, PzlNotifiedAt tom,
            │     og (PzlSeverity i ALWAYS_NOTIFY_SEVERITIES eller PzlNotifyTeams = ja)
            │     → sett PzlNotifiedAt (claim) → send → ved feil: nullstill PzlNotifiedAt
            └─ Nyheter: PromotedState = 2, FirstPublishedDate innen NEWS_LOOKBACK_DAYS,
                  publisert (hovedversjon, godkjent), ikke registrert i Table Storage
                  → registrer i tabellen (claim) → send → ved feil: fjern raden
```

- Mottakere: `transitiveMembers` på gruppene via Graph, dedupliserte, uten deaktiverte kontoer og (som standard) uten gjester.
- Sending: `POST /teamwork/sendActivityNotificationToRecipients` i grupper på 100, med retry ved 429/5xx.
- Nyhetssider skrives aldri til. En feltoppdatering på en publisert side ville gitt en ny versjon og «upubliserte endringer», så sendt-status for sider ligger i Table Storage i Function-appens lagringskonto.
- Timer-triggere kjører som singleton, og claim-før-send gjør at overlappende kjøringer uansett ikke dobbeltsender.

## Forutsetninger

- Pzl.Part.Announcement 1.1.0 eller nyere installert, med hovedmalen kjørt på nytt så de nye kolonnene finnes.
- Azure-abonnement, og rettigheter til å opprette ressurser og tildele app-roller i Entra (Global Administrator eller Privileged Role Administrator for steg 2).
- Verktøy: [Azure CLI](https://learn.microsoft.com/cli/azure/), [Azure Functions Core Tools](https://learn.microsoft.com/azure/azure-functions/functions-run-local) v4, Node 20 eller 22, PnP.PowerShell, Microsoft.Graph PowerShell.

## Steg 1: Azure-ressurser

`scripts/Deploy-NotificationInfrastructure.ps1` oppretter alt som trengs, idempotent (kan kjøres på nytt): ressursgruppe, user-assigned managed identity, lagringskonto, Application Insights, rolletildelinger for identiteten på lagringskontoen, Function-app på Flex Consumption (Node 22) med identitetsbasert tilgang til lagring, og app-innstillingene fra steg 5. Krever Azure CLI 2.70+ innlogget med rettigheter til å opprette ressurser og rolletildelinger.

```powershell
az login --tenant <tenant-id>
.\scripts\Deploy-NotificationInfrastructure.ps1 `
  -SubscriptionId <abonnement> `
  -ResourceGroupName rg-intranet-notification `
  -Location westeurope `
  -FunctionAppName func-intranet-notification `
  -StorageAccountName stintranetnotify01 `
  -ConfigSiteUrl https://<tenant>.sharepoint.com
```

Regionen må støtte Flex Consumption (`az functionapp list-flexconsumption-locations`). Norway East gjør det ikke per oktober 2026, West Europe gjør det. Valgfrie parametere: `-TeamsDeepLinkAppId`, `-TeamsDeepLinkEntityId`, `-ConfigListTitle`, `-AnnouncementListTitle`, `-AlwaysNotifySeverities`, `-DryRun` (standard `$true`).

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

1. Kjør hovedmalen på nytt på områdene med driftsmeldinger, så kolonnene `PzlNotifyTeams`, `PzlPinned` og `PzlNotifiedAt` finnes (`Invoke-PnPSiteTemplate -Path .\Templates\Announcement.xml`).
2. Opprett konfigurasjonslista på området som skal eie konfigurasjonen (typisk rot- eller hub-området):

   ```powershell
   Connect-PnPOnline -Url https://<tenant>.sharepoint.com -Interactive -ClientId <pnp-client-id>
   Invoke-PnPSiteTemplate -Path .\Templates\NotificationConfig.xml
   ```

3. Legg inn én rad per område i **Varslingskonfigurasjon**:

   | Kolonne | Verdi |
   |---|---|
   | Tittel | Visningsnavn, brukes i loggen |
   | Områdeadresse | Full URL til området |
   | Varsle nyheter | Ja/nei |
   | Varsle driftsmeldinger | Ja/nei |
   | Standard mottakergrupper | Objekt-ID for én eller flere Entra-grupper, én per linje. Brukes når elementet ikke har målgruppe. |
   | Driftsmeldingsliste | Tom for standard (`Driftsmeldinger`), ellers tittel på lista |
   | Aktiv | Ja |

   Et område med egen mottakergruppe (for eksempel en lederportal) får en egen rad med den gruppa som standard.

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

   Tekstene i `activityTypes` i `teams-app/manifest.json` («Driftsmelding: {title}», «Ny nyhet: {title}») er det brukeren ser i bjella og kan tilpasses før bygging. Linjen under varselet viser tittelen på konfigurasjonsraden i steg 3, så gi radene navn som gir mening for brukeren («Svaneposten», «Lederportalen»).

2. Last opp zip-fila i Teams admin center (Teams apps > Manage apps > Upload new app).
3. Lag en **app setup policy** som installerer appen for mottakerne (uten å feste den i appfeltet), og tildel den til samme grupper som i konfigurasjonen. Start med en testgruppe.
4. Pek varslene på en Teams-fane. Graph godtar bare Teams-deep-linker som `webUrl`, så dette er påkrevd. Enklest er varsel-appens egen fane, som åpner intranettet:

   ```bash
   az functionapp config appsettings set -g <ressursgruppe> -n <function-app> --settings TEAMS_DEEPLINK_APP_ID=<Teams-app-ID fra punkt 1> TEAMS_DEEPLINK_ENTITY_ID=intranet
   ```

   Fanen i varsel-appen sender `{subEntityId}` videre til intranettet som `?driftsmelding=`, og det er den parameteren webdelen bruker for å åpne riktig driftsmelding.

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
| `ALWAYS_NOTIFY_SEVERITIES` | `Varsel,Alert` | Meldingsnivåer som alltid varsles |
| `NEWS_LOOKBACK_DAYS` | `7` | Hvor langt tilbake nye nyheter hentes |
| `INCLUDE_GUESTS` | `false` | Ta med gjestebrukere som mottakere |
| `TEAMS_DEEPLINK_APP_ID` | *(påkrevd)* | Teams-app som varslene åpner: varsel-appen selv (steg 4) eller intranett-appen. Uten denne sendes ingenting, og loggen sier hvorfor. |
| `TEAMS_DEEPLINK_ENTITY_ID` | `intranet` | `entityId` for fanen i den appen. `intranet` er varsel-appens fane. |
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
Invoke-RestMethod -Method Post -Uri "https://func-intranet-notification.azurewebsites.net/api/test-notification?userId=$userId&type=announcement" -Headers @{ 'x-functions-key' = $key }
```

Svaret inneholder `ok` og eventuell Graph-feil. Varselet skal dukke opp i bjella innen sekunder. `type=news` tester nyhetsvarianten. Legg til `&bulk=1` for å sende via masse-endepunktet.

> **Malparametere må stemme nøyaktig** med plassholderne i `templateText` i manifestet (`{title}`). Masse-endepunktet godtar ekstra parametere, men leverer da ingenting, uten feil. Per-bruker-endepunktet svarer «Incorrect template parameter arity». Endrer du `templateText`, må `templateParameters` i `announcements.ts`/`news.ts` endres tilsvarende.


1. Konfigurasjonsrad mot en testgruppe, setup policy mot samme gruppe.
2. Test de tre variantene: nyhet, driftsmelding med nivå «Varsel», driftsmelding med «Varsle i Teams».
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
| `401`/`403` fra SharePoint | App-rollen `Sites.Selected` på **SharePoint Online** (ikke bare Graph), og `Grant-PnPAzureADAppSitePermission` på riktig område med riktig nivå |
| `403` fra `transitiveMembers` | `GroupMember.Read.All` tildelt, og at admin consent er gitt (scriptet tildeler app-roller direkte, consent er implisitt) |
| `400` «Invalid 'webUrl' specified» fra `sendActivityNotificationToRecipients` | `TEAMS_DEEPLINK_APP_ID` mangler eller er ikke en Teams-app-ID. Lenken må være `https://teams.microsoft.com/l/...` |
| `400`/`403` ellers fra `sendActivityNotificationToRecipients` | `TeamsActivity.Send`, at Teams-appen er installert for mottakeren (setup policy), og at `webApplicationInfo.id` i manifestet er client ID til identiteten som sender |
| Ingen mottakere | Gruppe-ID-ene i konfigurasjonen, og at gruppa har medlemmer som ikke er deaktivert/gjester |
| Loggen sier «Sent … to N recipients», men ingenting i bjella | Appen var ikke installert for mottakeren da det ble sendt (setup policy tar tid), eller `templateParameters` matcher ikke `templateText`. Send et testvarsel med `/api/test-notification` til én bruker for å få feilen. |
| Nyheter varsles ikke | `PromotedState = 2`, publisert hovedversjon, innenfor `NEWS_LOOKBACK_DAYS`, og ikke allerede en rad i tabellen `NotificationState` |
| Driftsmeldinger varsles ikke | Startdato passert, sluttdato ikke passert, `PzlNotifiedAt` tom, og nivå/«Varsle i Teams» |
| Dobbelt varsel | Viva Connections kan også varsle om nyheter. Skru av det ene. |

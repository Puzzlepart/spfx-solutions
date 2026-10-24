# Teams-varsling for nyheter og driftsmeldinger

Valgfritt tillegg til [Pzl.Part.Announcement](../README.md). En timer-trigget Azure Function sender varsler i Teams-aktivitetsfeeden (bjella) når

- en **nyhetsside** publiseres på et konfigurert område, og
- en **driftsmelding** blir aktiv og enten har meldingsnivået «Varsel» eller har «Varsle i Teams» satt.

Varslene respekterer **målgruppestyring**: har elementet målgrupper, varsles medlemmene i de gruppene. Ellers varsles standardgruppene for området. Klikk på et varsel åpner intranett-fanen i Teams med elementet i fokus (driftsmeldinger åpnes i popover av webdelen), eller SharePoint-siden direkte hvis ingen Teams-fane er konfigurert.

Lista og webdelen fungerer uten dette tillegget. Kolonnene `PzlNotifyTeams`, `PzlPinned` og `PzlNotifiedAt` finnes i hovedmalen fra versjon 1.1.0 og er uten effekt til tjenesten tas i bruk.

## Innhold

| Mappe | Hva |
|---|---|
| `function/` | Azure Function (TypeScript, Node 20+, Functions v4). Kjører hvert minutt. |
| `teams-app/` | Manifest-mal og ikoner for Teams-appen som eier varslene. |
| `scripts/` | `Grant-NotificationPermissions.ps1`: tildeler app-rettigheter til tjenesteidentiteten. |
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

Alle navn og regioner kan tilpasses. Eksempelet bruker Flex Consumption-plan og en user-assigned managed identity.

```bash
RG=rg-intranet-notification
LOC=norwayeast
STORAGE=stintranetnotify$RANDOM
FUNC=func-intranet-notification
MI=mi-intranet-notification

az group create -n $RG -l $LOC
az storage account create -n $STORAGE -g $RG -l $LOC --sku Standard_LRS --allow-blob-public-access false
az identity create -n $MI -g $RG -l $LOC
az monitor app-insights component create --app ai-intranet-notification -g $RG -l $LOC --kind web

az functionapp create -n $FUNC -g $RG --storage-account $STORAGE \
  --flexconsumption-location $LOC --runtime node --runtime-version 22 \
  --assign-identity $(az identity show -n $MI -g $RG --query id -o tsv) \
  --app-insights ai-intranet-notification
```

Gi identiteten tilgang til lagringskontoen (Functions-runtime og state-tabellen bruker identitet, ikke nøkler):

```bash
MI_PRINCIPAL=$(az identity show -n $MI -g $RG --query principalId -o tsv)
MI_CLIENT=$(az identity show -n $MI -g $RG --query clientId -o tsv)
STORAGE_ID=$(az storage account show -n $STORAGE -g $RG --query id -o tsv)
for ROLE in "Storage Blob Data Owner" "Storage Queue Data Contributor" "Storage Table Data Contributor"; do
  az role assignment create --assignee-object-id $MI_PRINCIPAL --assignee-principal-type ServicePrincipal --role "$ROLE" --scope $STORAGE_ID
done
az functionapp config appsettings set -n $FUNC -g $RG --settings \
  AzureWebJobsStorage__accountName=$STORAGE \
  AzureWebJobsStorage__clientId=$MI_CLIENT \
  AzureWebJobsStorage__credential=managedidentity \
  AZURE_CLIENT_ID=$MI_CLIENT
az functionapp config appsettings delete -n $FUNC -g $RG --setting-names AzureWebJobsStorage
```

Noter `MI_CLIENT` (client ID) og `MI_PRINCIPAL` (object ID). De brukes i de neste stegene.

## Steg 2: Rettigheter i Entra

Tjenesteidentiteten trenger disse applikasjonstillatelsene:

| Ressurs | Tillatelse | Brukes til |
|---|---|---|
| Microsoft Graph | `TeamsActivity.Send` | Sende aktivitetsvarsler |
| Microsoft Graph | `GroupMember.Read.All` | Finne medlemmer i mottakergruppene |
| Microsoft Graph | `Sites.Selected` | Kreves sammen med SharePoint-rollen under |
| SharePoint Online | `Sites.Selected` | Lese lister via SharePoint REST, skrive `PzlNotifiedAt` |

```powershell
.\scripts\Grant-NotificationPermissions.ps1 -ServicePrincipalObjectId <MI_PRINCIPAL>
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

1. Kopier `teams-app/` og fyll inn plassholderne i `manifest.json`:

   | Plassholder | Verdi |
   |---|---|
   | `{{TEAMS_APP_ID}}` | Ny GUID (`[guid]::NewGuid()`) |
   | `{{NOTIFICATION_CLIENT_ID}}` | `MI_CLIENT` |
   | `{{SITE_URL}}` / `{{SITE_URL_ENCODED}}` | Intranettets URL, og samme URL URL-kodet |
   | `{{SITE_HOST}}` | `<tenant>.sharepoint.com` |
   | `{{APP_NAME}}`, `{{DEVELOPER_NAME}}` | Visningsnavn og organisasjon |

   Bytt gjerne ikonene, så appen får samme ikon som intranett-appen. Tekstene i `activityTypes` er det brukeren ser i bjella og kan oversettes.

2. Zip `manifest.json`, `color.png` og `outline.png` og last opp i Teams admin center (Teams apps > Manage apps > Upload new app).
3. Lag en **app setup policy** som installerer appen for mottakerne (uten å feste den i appfeltet), og tildel den til samme grupper som i konfigurasjonen. Start med en testgruppe.
4. Finn Teams-app-ID og `entityId` for intranett-fanen som varslene skal åpne. Legg dem i `TEAMS_DEEPLINK_APP_ID` og `TEAMS_DEEPLINK_ENTITY_ID` i steg 5. Uten disse lenker varslene rett til SharePoint-siden.

> **Ikke verifisert i alle tenanter:** at en managed identity aksepteres som `webApplicationInfo.id`. Test tidlig med én mottaker. Alternativet er en app-registrering med managed identity som federated identity credential; Function-koden er den samme (`AZURE_CLIENT_ID` settes til app-registreringens client ID).

## Steg 5: Deploy Function-appen

```bash
cd notification/function
npm install
npm run build
func azure functionapp publish $FUNC
```

Sett app-innstillingene:

```bash
az functionapp config appsettings set -n $FUNC -g $RG --settings \
  CONFIG_SITE_URL=https://<tenant>.sharepoint.com \
  TEAMS_DEEPLINK_APP_ID=<intranett-appens id> \
  TEAMS_DEEPLINK_ENTITY_ID=<fanens entityId> \
  DRY_RUN=true
```

| Innstilling | Standard | Beskrivelse |
|---|---|---|
| `AZURE_CLIENT_ID` | | Client ID til managed identity (settes i steg 1) |
| `CONFIG_SITE_URL` | *(påkrevd)* | Område med konfigurasjonslista |
| `CONFIG_LIST_TITLE` | `Varslingskonfigurasjon` | Tittel på konfigurasjonslista |
| `ANNOUNCEMENT_LIST_TITLE` | `Driftsmeldinger` | Standard tittel på driftsmeldingslista |
| `ALWAYS_NOTIFY_SEVERITIES` | `Varsel,Alert` | Meldingsnivåer som alltid varsles |
| `NEWS_LOOKBACK_DAYS` | `7` | Hvor langt tilbake nye nyheter hentes |
| `INCLUDE_GUESTS` | `false` | Ta med gjestebrukere som mottakere |
| `TEAMS_DEEPLINK_APP_ID` | | Teams-app som varslene skal åpne (intranett-appen) |
| `TEAMS_DEEPLINK_ENTITY_ID` | | `entityId` for fanen i den appen |
| `STATE_STORAGE_ACCOUNT_NAME` | Function-appens konto | Lagringskonto for sendt-status for sider |
| `STATE_TABLE_NAME` | `NotificationState` | Tabellnavn |
| `TIMER_SCHEDULE` | `0 * * * * *` | NCRONTAB, hvert minutt |
| `DRY_RUN` | `false` | Logg hva som ville blitt sendt, uten å sende eller skrive |

Start med `DRY_RUN=true` og se i loggen (Application Insights eller `func azure functionapp logstream`) at riktige elementer og mottakerantall plukkes opp. Sett så `DRY_RUN=false`.

## Steg 6: Test og pilot

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
| `400`/`403` fra `sendActivityNotificationToRecipients` | `TeamsActivity.Send`, at Teams-appen er installert for mottakeren (setup policy), og at `webApplicationInfo.id` i manifestet er client ID til identiteten som sender |
| Ingen mottakere | Gruppe-ID-ene i konfigurasjonen, og at gruppa har medlemmer som ikke er deaktivert/gjester |
| Nyheter varsles ikke | `PromotedState = 2`, publisert hovedversjon, innenfor `NEWS_LOOKBACK_DAYS`, og ikke allerede en rad i tabellen `NotificationState` |
| Driftsmeldinger varsles ikke | Startdato passert, sluttdato ikke passert, `PzlNotifiedAt` tom, og nivå/«Varsle i Teams» |
| Dobbelt varsel | Viva Connections kan også varsle om nyheter. Skru av det ene. |

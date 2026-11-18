# Driftsmeldinger – domenebegreper

Felles språk for `Pzl.Part.Announcement` (webdel og banner) og tillegget `Pzl.Ext.Notification` (Teams-varsling). Bruk disse begrepene i kode, README, feltbeskrivelser og samtaler.

## Begreper

**Driftsmelding**
Et element i lista *Driftsmeldinger*: en beskjed til brukerne om at noe er, var eller blir påvirket i IT- eller driftstjenestene (feil, nedetid, planlagt vedlikehold). Generelle kunngjøringer (HR, kommunikasjon) er ikke driftsmeldinger. Til det brukes SharePoint-nyheter. *Announcement* er bare det tekniske navnet i koden.

**Sentral liste**
Én liste *Driftsmeldinger* per tenant, typisk på intranettets rotområde. Webdelen og banneret leser fra den, og andre områder peker på den med «Område for driftsmeldinger» (`listSiteUrl`). Lister per område er et unntak som bare varslingstjenesten trenger å kjenne til.

**Meldingsnivå** (`PzlSeverity`)
Hvor alvorlig driftsmeldingen er for brukerne: *Informasjon*, *Advarsel* eller *Kritisk*. Nivået gir farge og ikon. Standard er *Advarsel*. Eldre verdier tolkes slik: *Varsel* → *Kritisk*, *Normal* → *Informasjon*. Verdiene tolkes med en fast tabell, ikke med brukerens språk.

**Status** (`PzlStatus`)
Hvor langt hendelsen har kommet: *Planlagt*, *Pågår* eller *Løst*. Feltet er påkrevd, og standard er *Planlagt*. Tom status (eldre elementer) tolkes som *Pågår*. *Løst* overstyrer nivåets farge (grønn, med hake).

**Visningsperiode** (`PzlStartDate`–`PzlEndDate`)
Tidsrommet driftsmeldingen vises på intranettet. Det sier ikke når hendelsen skjer. For planlagt vedlikehold står tidspunktet for arbeidet i teksten.

**Festet** (`PzlPinned`)
Driftsmeldingen sorteres øverst, foran status og nivå.

**Banner** (`PzlShowAsBanner`)
Driftsmeldingen vises også øverst på sidene, via customizeren *AnnouncementBanner*. Valget gjøres manuelt per melding. I webdelens skjema er det forhåndsvalgt når nivået er *Kritisk*. Brukeren kan lukke et banner, men det kommer tilbake når meldingen endres. Det er slik brukerne får se at en melding er satt til *Løst*.

**Hurtighandling**
Knappene «Marker som løst», «Marker som pågår» og «Rediger» under «For redaktører» i detaljdialogen. De vises for brukere med redigeringsrett på den sentrale lista. «Rediger» åpner driftsmeldingen i webdelens eget skjema (samme som for ny driftsmelding), ikke i listeskjemaet.

**Detaljdialog**
Modaldialogen midt på siden som åpnes ved klikk på en driftsmelding. Den er lik enten den åpnes fra webdelen eller fra banneret, og viser feltene som er valgt under «Felter i detaljvisning». Det finnes ingen popover lenger.

**For redaktører**
Sammenslått seksjon nederst i detaljdialogen, bare synlig for brukere med redigeringsrett på den sentrale lista. Her ligger hurtighandlingene og varslingsvalgene.

**Varsling** (tillegg)
Et Teams-aktivitetsvarsel (bjella) som sendes av `Pzl.Ext.Notification`. Varsling er en leveransekanal, ikke en egen meldingstype. Driftsmeldinger fungerer fullt ut uten den. Feltene `PzlNotifyTeams` («Varsle i Teams»), `PzlNotifyAgain` («Varsle på nytt») og `PzlNotifiedAt` («Varslet i Teams») hører til tillegget og legges til lista med tilleggets mal. Redaktører setter «Varsle i Teams» i skjemaet for ny driftsmelding og i detaljdialogen, og «Varsle på nytt» i detaljdialogen. Endringer i detaljdialogen lagres med en gang.

**Nivåer som alltid varsles** (`PzlNotifySeverities`)
Meldingsnivåene tjenesten varsler om uansett hva «Varsle i Teams» står til, satt per varslingskilde. Webdelen leser dem fra *Varslingskonfigurasjon* og viser «Varsle i Teams» som valgt og låst for disse nivåene.

**Varslingskilde**
En rad i *Varslingskonfigurasjon*: et område med driftsmeldinger og/eller nyheter som tjenesten skal varsle om, med egne mottakergrupper. Varslingen støtter flere kilder, uavhengig av den sentrale lista.

## Sortering i webdelen

1. Festet
2. Status: *Pågår* → *Planlagt* → *Løst*
3. Nivå: *Kritisk* → *Advarsel* → *Informasjon*
4. Nyeste startdato først

## Varslingsregler (tillegget)

- En driftsmelding varsles første gang når den blir aktiv og har «Varsle i Teams» eller et nivå i «Nivåer som alltid varsles».
- En driftsmelding som er *Løst* når den blir aktiv, varsles ikke.
- En driftsmelding som fantes og var aktiv før varslingen ble slått på for området, får ikke første varsel. Bare nye varsles: opprettet etter starten, eller med startdato etter starten.
- Tittelen i varselet: *Pågår* → «Driftsmelding: <tittel>», *Planlagt* → «Planlagt: <tittel>».
- «Varsle på nytt» sender et nytt varsel («Oppdatert: <tittel>», eller «Løst: <tittel>» når status er *Løst*). Tjenesten setter feltet tilbake til *Nei*. Det finnes ingen automatisk ny varsling.

# Plan: Driftsmeldinger 2.0

Beslutningene og begrepene står i [CONTEXT.md](CONTEXT.md). Arbeidet er gjort i fire faser med én commit per fase.

**Status:** alle fire fasene er gjennomført. Bygging, pakking og enhetstester er grønne. Dette gjenstår:

- Kjør `scripts/Install-Driftsmeldinger.ps1` mot en testtenant, først med `-WhatIf`, både som nyinstallasjon og som oppgradering fra 1.5.4. Skriptet er syntakssjekket, men ikke kjørt mot SharePoint. Det samme gjelder `Pzl.Ext.Notification/scripts/Install-Notification.ps1` (Azure, Entra og Teams).
- Test i en tenant: oppgradering fra 1.5.4 (malen kjørt på nytt på en liste med eksisterende elementer), hurtighandlinger, «Område for driftsmeldinger» mot et annet område, og banneret med *Løst*.
- `Pzl.Ext.Notification/Templates/AnnouncementFields.xml` gjentar alle feltreferansene i innholdstypen. Bekreft at PnP legger til de tre nye feltene uten å endre resten.
- `mpnId` i `package-solution.json` er satt til SPFx-standarden (`Undefined-1.24.1`). Erstatt den med SoftwareOnes Partner ID hvis den skal vises.

## Fase 1: Skill ut `Pzl.Ext.Notification`

- Flytt `notification/` (Azure-funksjon, Teams-app, skript) og `Templates/NotificationConfig.xml` med objekter til en ny løsning, `Pzl.Ext.Notification`.
- Flytt customizeren *AnnouncementDeepLink* til et eget SPFx-prosjekt i den nye løsningen, med ny løsnings-ID, feature-ID og komponent-ID.
- Fjern håndteringen av `page:` fra webdelen. Webdelen beholder `?driftsmelding=<id>` og `driftsmelding:<id>`.
- Fjern feature og `elements.xml` for deep link fra `Pzl.Part.Announcement`.

## Fase 2: Datamodell

- Meldingsnivå: *Informasjon / Advarsel / Kritisk*.
- Nytt felt: *Status* (`PzlStatus`): *Planlagt / Pågår / Løst*, påkrevd, standard *Planlagt*.
- Feltbeskrivelser for alle feltene.
- Fast verditabell for nivå og status i koden (uavhengig av brukerens språk).
- Teams-feltene (`PzlNotifyTeams`, `PzlNotifiedAt`) fjernes fra hovedmalen og flyttes til en mal i `Pzl.Ext.Notification`, som også legger til `PzlNotifyAgain`.
- Varslingstjenesten: status, «Varsle på nytt», nye titler, nye nivåer i konfigurasjonen.

## Fase 3: Webdel og banner

- Egenskapen «Område for driftsmeldinger» (`listSiteUrl`) på webdelen. Rettighetene sjekkes på den sentrale lista.
- Sortering og visning etter status (*Løst* grønn med hake, *Planlagt* med kalenderikon og etikett).
- Hurtighandlinger i popoveren for brukere med redigeringsrett.
- «Vis som banner» forhåndsvalgt for *Kritisk* i skjemaet.
- Banneret viser *Løst* grønt og vises på nytt etter endring (versjonsnøkkelen er `Modified`).

## Fase 4: Pakke og dokumentasjon

- Visningsnavnet blir «Driftsmeldinger», og utvikler blir SoftwareOne.
- Versjon 2.0.0. Løsnings-ID og komponent-ID-er beholdes.
- README og oppgraderingsnotater for dem som bruker lista og webdelen i dag.

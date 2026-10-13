# Microsoft Graph: einfaches Mehrbenutzerkonzept

**Stand: 12.10.2026, Europe/Berlin. Akzeptierter Umsetzungsvertrag.**

Der isolierte 4.0-Kandidat implementiert die persönlichen Konto-/UI-,
Sessionbindungs- und Workerpfade. Implementierungsstatus, konservative
Host-Fallbacks und Nachweise: [MULTI_USER.md](MULTI_USER.md).
Live-Aktivierung und echte Mehrbenutzer-OAuth-Canaries sind nicht erfolgt.
Geprüfter Host: OpenClaw 2026.10.1 (e178c39); Pluginbasis: 3.12.1.
Ziel: eine persönliche OAuth-Verbindung je verifiziertem OpenClaw-Benutzer,
unveränderte bestehende Graph-Tools und dokumentierte OpenClaw-Betriebsarten.
Keine Core-Änderungen, keine privaten Core-Imports, keine Live-Aktivierung/Migration.

## 1. Entscheidung und Problemrahmen

Nicht versuchen, bei jedem Tool-Aufruf den zuletzt sprechenden Menschen neu zu
ermitteln. Stattdessen **OAuth-Eigentum, Ausführungskonto und Berechtigung trennen**.
Ein Ausführungskontext hat genau ein gebundenes Graph-Konto. Die gleichen `me`-
Tools arbeiten mit diesem Konto; ein weiterer Teilnehmer ändert es nicht implizit.

Das ist eine neue, ausdrücklich dokumentierte Pluginregel. Sie ersetzt für
gemeinsam bediente Sessions den früher angenommenen Vertrag „jedes me folgt
immer dem aktuell sprechenden Menschen“. Das ursprüngliche Gesamtgoal ist damit
nicht bereits erfüllt; die geänderte Auswahlsemantik muss in seine Umsetzung
übernommen und getestet werden.

| Alternative | Vorteil | Entscheidender Nachteil | Entscheidung |
| --- | --- | --- | --- |
| Konto bei jedem Aufruf aus aktuellem Sprecher | Intuitiv in persönlichen Chats | Live Sprecheridentität bei Steering/Autonomie nicht überall öffentlich bewiesen; Ziel wechselt mitten im Auftrag | Nicht als allgemeiner Vertrag |
| Sessionfestes Konto plus explizite Agentrechte | Gleiche Tools, stabile Ziele, Teamkollaboration möglich | Gemeinsame Bedienung und Datenanzeige müssen bewusst freigegeben sein | Empfohlen |
| Nur ein festes Konto je Agent | Sehr einfache autonome Ausführung | Mehr Konten brauchen zusätzliche Agenten; beweist keinen menschlichen Auftraggeber | Kleine Ergänzung für autonome Worker |

Kreative Auflösung des Konflikts: Identität **bei der Kontobindung** prüfen,
Ausführungsrechte **bei jedem Zugriff**. Für autonome Arbeit ausdrücklich
erteilte Dauerautorität verwenden, statt eine anwesende Person zu erfinden.

## 2. Was OpenClaw vorgibt — und was nicht

- Ein gemeinsam bedienter Agent gehört zu einer Vertrauensgemeinschaft. Für
  gegeneinander isolierte Personen sind getrennte Agenten/Gateway-Grenzen nötig.
  [Multi-user mode](https://docs.openclaw.ai/concepts/multi-user#trust-boundary),
  [Gateway trust model](https://docs.openclaw.ai/gateway/security/trust-model).
- Session-Eigentum ist Zuständigkeit, keine Übertragung persönlicher Kontorechte.
  `shared` ist der reguläre UI-Startmodus; andere Modi steuern Kollaboration.
  [Session-Konfiguration](https://docs.openclaw.ai/gateway/config-agents/sessions).
- Persönliche **Modellkonten** liefern ein nützliches Vorbild: Auswahl bleibt
  sessionfest; Mitsteuerung und Forks behalten sie; eine neue persönliche Auswahl
  verlangt die Identität des Auswahlenden. Das ist keine vorhandene Graph-API.
  [Kontopins](https://docs.openclaw.ai/concepts/multi-user#pin-and-default-rules).
- Geplante Läufe erben nicht den Requestzustand ihres Erstellungsdialogs.
  Das ist ein dokumentiertes Autonomiemodell, nicht für sich eine Sicherheitslücke.
  [Automation-Laufzeit](https://docs.openclaw.ai/automation/cron-jobs/how-it-works).
- Subagent-Modellauth wird nach Agent aufgelöst. Daraus folgt **keine** automatische
  Graph-OAuth-Vererbung. [Subagents](https://docs.openclaw.ai/tools/subagents/nesting).

**Wichtige Korrektur:** Gemeinsamer Gesprächskontext ist nicht automatisch eine
Freigabe sämtlicher OAuth-Konten. Gegenseitige Graph-Nutzung ist möglich, sobald
das Plugin sie ausdrücklich freigibt. Selbst dann gelten die Kontorechte des
Agents, Microsoft-Scopes und vorhandenen Ressourcen-/Operationsgrenzen.

## 3. Vier kleine Bausteine

1. **Verbindung:** kanonische OpenClaw-Profil-ID → maximal eine aktive persönliche
   Microsoft-Verbindung. Keine eigene Benutzerverwaltung. Nicht verbunden ist
   ein zulässiger Zustand; „genau eine“ heißt nicht Zwangsanmeldung für alle.
2. **Kontoberechtigung:** bestehende Policy pro Verbindung, statt global:
   **Konto × Agent × Dienst × Operation/Ressource**. Default deny bleibt.
3. **Sessionbindung:** pluginintern verwaltete Referenz auf genau eine Verbindung,
   mit Sessiongeneration und Bindungsrevision. UI zeigt das aktive Konto.
4. **Optionaler Worker:** vertraute Host-Agent-ID → genau eine Verbindung,
   ausdrücklich für delegierte/unbeaufsichtigte Nutzung freigegeben.

Keine Benutzer-ID in den vorhandenen `me`-Toolargumenten. Keine Auswahl aus
Anzeigenamen, E-Mail, Prompt, Session-Assignee oder zuletzt benutztem Token.
Eine UI-/Modellauswahl eines Kontos ist ein Zielwunsch, niemals selbst Autorität.

**Eindeutige Kontoauflösung:** Bei einem kontogebundenen Worker bestimmt dessen
feste Bindung das Ziel unveränderlich. Persönliche automatische Erstbindung und
UI-Kontowechsel sind dort nicht zulässig. Eine widersprechende Sessionbindung
führt zu `graph_execution_account_conflict`, nicht zu Übersteuerung oder Fallback.
Bei normalen Agents gilt die geschützte Sessionbindung; nur eine ungebundene
persönliche Session darf authentifiziert initialisiert werden. Ohne gültige
Bindung gibt es keinen Graph-Zugriff. Ein Worker braucht somit keine zweite
änderbare Sessionauswahl.

## 4. Persönliche und gemeinsame Sessions

### Persönliche Nutzung

Eine neue persönliche Session erhält die Verbindung des verifizierten
initiierenden Benutzers. Die automatische Initialisierung muss an einer
unterstützten authentifizierten Admission erfolgen; alternativ „Mein Konto
verwenden“ im authentifizierten Pluginpanel. Eine nicht beweisbare Initialisierung
bleibt ungebunden, statt Eigentum zu raten. DM-Kanäle benötigen verifizierte
Profilverknüpfung, nicht nur einen numerischen Sender. Session-Isolation folgt
der dokumentierten DM-Konfiguration, etwa `per-channel-peer`.

Persönliche Nutzung verspricht keine Geheimhaltung gegenüber Gateway-Admins,
Hostcode oder allen Werkzeugen eines gemeinsam bedienten Agents. Gegenüber
Mitbenutzern gewünschte Isolation erfordert die entsprechende OpenClaw-Grenze.

### Gemeinsame Nutzung

Die Kontoverbindung hat eine explizite **Freigabe für gemeinsam bediente Agents/
Sessions**, zusätzlich zu ihren Agent-/Dienstrechten. Ohne diese Freigabe darf
ein bloßes `shared` keine persönliche Graph-Ausführung aktivieren. OAuth-Anmeldung
allein erteilt keine Teamfreigabe. Die Freigabe erfolgt durch den Kontoinhaber;
Administration darf die Policy einschränken und genehmigte Migration ausführen,
aber fremde Kontofreigabe nicht stillschweigend behaupten.

Danach können alle durch OpenClaw zugelassenen Mitbediener die freigegebenen
Fähigkeiten dieses Agents im gebundenen Konto auslösen. Alice kann Bobs Konto
nutzen, **wenn** Bob diese Nutzung und den ausführenden Agent freigegeben hat.
Das Konto wechselt nicht automatisch mit dem Sprecher. Ein bewusstes Wechseln
ist nur zu einem dafür freigegebenen Konto zulässig und nur zwischen Aufträgen.

UI-Warntext bei Freigabe und dauerhaft am Konto-/Berechtigungsbereich:

> Gemeinsam bedienbare Agents können die hier freigegebenen Microsoft-Aktionen
> in diesem Konto auch auf Anweisung anderer zugelassener Benutzer ausführen.
> Leser gemeinsamer Sessions können dort ausgegebene Microsoft-Daten sehen.
> OAuth-Tokens und die Verwaltung deiner Verbindung werden dadurch nicht geteilt.

Read-only/Suggest/Public-Sichtbarkeit kann Daten offenlegen, ohne neue Aktionen
zu erlauben. Auch Rückgaben delegierter Arbeit können in der Elternsession
sichtbar werden. Deshalb gilt die Ausgabe-/Liefergrenze zusätzlich zur
Bedienberechtigung; private Ergebnisse dürfen nicht durch Delegation in eine
nicht freigegebene gemeinsame/public Destination gelangen. Bereits ausgegebene
Daten verschwinden bei späterem Widerruf nicht.

Die Pluginfreigabe ist bewusst keine komplexe zweite Teilnehmer-ACL. OpenClaw
regelt, wer Agent/Session bedienen darf; das Plugin regelt die Graph-Fähigkeiten.
Fehlt eine belastbare Hostfeststellung des erlaubten Kontextes, keine private
Isolation behaupten: explizite Teamfreigabe verlangen oder Zugriff ablehnen.

## 5. Delegation und Automationen ohne erfundene Identität

**Für die erste Umsetzung: kontogebundene Worker als minimale robuste Lösung.**
Beispiel: `graph-robin` ist serverseitig Robins Verbindung zugeordnet und erhält
nur die nötigen Dienste/Operationen. Ein Job oder eine Delegation richtet sich an
diesen konfigurierten Agent. Damit ist das Zielkonto deterministisch; der Agent
muss nicht bei jedem Aufruf einen angeblich aktuellen Benutzer nennen.

Das Konto muss delegierte/unbeaufsichtigte Nutzung ausdrücklich erlauben.
Wer den Worker nach OpenClaw-Policy starten oder steuern darf, kann dessen
Graph-Fähigkeiten auslösen. Die Start-/Delegationsrechte gehören deshalb zur
Freigabeentscheidung. Ein fremd startbarer Worker ist keine private Benutzergrenze.
Für private Worker müssen die OpenClaw-Routing-/Agent-/Gateway-Grenzen passen.

„Erstellt von“, „für Benutzer“ und „ausgeführt mit Konto“ bleiben getrennte
Attributionen. Der kanonische Kontoinhaber ist aus der administrativ/inhaberseitig
erteilten Workerbindung bekannt; die menschliche Erstellerattribution kommt
nur aus Hostbelegen. Fehlt sie, steht dort „nicht nachgewiesen“, nicht ein
modellgenannter Benutzer.

Der Scheduler bleibt OpenClaw; das Plugin baut keinen zweiten Scheduler.
Jobziel `agentId` ist die einfache dauerhafte Auswahl. Der Jobprompt darf
„für Robin“ erklären, aber weder Binding noch Rechte erzeugen. Ein Job kann
nur erlaubte Worker wählen; Änderungen am Ziel sind eine neue Entscheidung.
Kontogeneration und aktuelle Freigaben werden auch nach Neustart geprüft.

**Spätere Vereinfachung bei vielen Konten:** eine account-scoped Delegation /
Jobbindung, die das Plugin vor dem Start serverseitig erzeugt und deren Rechte
höchstens die Eltern-/Dauerfreigabe sind. Kein Prompttoken als Autorisierung.
Standard-Spawns/Forks erhalten bis zum Nachweis öffentlicher Bindungsweitergabe
keine implizite Graph-Autorität. Sie nutzen einen freigegebenen Worker oder
melden `graph_execution_account_required`. Das ist keine pauschale Ablehnung
von Delegation, sondern der explizite unterstützte Aufrufpfad.

Ein unattended Job kann notwendige interaktive Schreibfreigaben nicht ersetzen.
Vorhandene Approval-Gates bleiben; fehlende zulässige Freigabe stoppt die Aktion.

## 6. Einheitliche Durchsetzung

Vor jedem Zugriff gilt:

`gültige Ausführungsbindung UND erlaubter Nutzungskontext UND aktuelle Konto-
Agent-Dienst-Operations-Ressourcenrechte UND OpenClaw-Toolpolicy UND OAuth-Scope
UND gegebenenfalls native Aktionsfreigabe`.

Erst danach Tokens lesen/refreshen und Graph aufrufen. Ein Benutzer ohne
Verbindung darf niemals auf das bisherige Robin-/Legacy-Konto zurückfallen.
Keine automatischen Konto-Fallbacks bei Ablauf oder Fehlern.

- Login/Status/Disconnect/Reconnect sind eigene authentifizierte Personen-APIs.
  Zugriff auf einen Agent oder ein Konto ist kein Recht auf Tokenexport,
  fremden Login oder Verbindungsverwaltung. Auch Adminansichten zeigen keine Tokens.
- Tokens, Refreshlocks und Caches sind verbindungs-/generationsgebunden.
  Kontowechsel oder Disconnect invalidiert zugehörige Approvals, Fortsetzungen,
  Uploadhandles und OneDrive-Instructions-Acknowledgements.
- Aktionsfreigaben binden Konto, Generation, Agent, Session, Call und effektive
  Parameter. Keine Übertragung einer Freigabe zwischen Konten.
- Kontowechsel während aktiver/zugelassener Arbeit ablehnen; nicht nur ein globales
  veränderliches `currentAccount` führen. Parallele Sessions bleiben unabhängig.
- Neue Zugriffe prüfen Widerruf aktuell; externe bereits angenommene Aktionen
  sind nicht rückholbar. Cleanup darf keine neue Autorität schaffen.
- Plugin-Sessionextension zeigt nur bereinigten Status. Ein generisch patchbarer
  Sessionwert ist **keine** autoritative Kontobindung. Bindungen kommen aus
  geschütztem Pluginzustand bzw. werden gegen ihn exakt validiert.

## 7. Öffentliche Schnittstellen und verbleibender Nachweis

| Mechanik | Evidenz am 12.10.2026 | Implementierungsfolge |
| --- | --- | --- |
| Gateway-Handler, Session-Zugriffsprüfung, aktuelle Autoritätsguards | `runtime-supported`: offizielle Runtime-Doku und installierte öffentliche SDK-Typen | Authentifizierte Konto-/Bindungs-APIs; vor Effekten revalidieren |
| Plugin-Sessionextension, Controls, Sessionreader | `runtime-supported`: offizielle Hooks-/Runtime-Doku und SDK-Typen | Statusprojektion/Lebenszyklus; nicht allein Autorisierung |
| Plugin-Subagent-Runner, exakte Agent-/Sessionidentität | `runtime-supported`: öffentliche Background-Work-Doku und SDK | Freigegebene Worker ohne privaten Core-Import |
| Automatische persönliche Erstbindung auf jedem DM-/Harnesspfad | `unconfirmed` als Ende-zu-Ende-Vertrag | Pro unterstütztlichem Pfad Canary; sonst authentifizierte UI-Bindung |
| Generische Graph-Bindungsweitergabe bei Spawn/Fork/Automationen | `unconfirmed`; Modellaccount-Vererbung beweist sie nicht | Erste Version: explizite Workerbindung, keine erfundene Vererbung |
| Genaues Sicherheitsverhalten des fertigen Plugins | Nicht implementiert oder E2E-geprüft | Tests und Review vor Release; keine Produktions-PASS-Aussage |

Quellen: [Gateway helpers](https://docs.openclaw.ai/plugins/sdk-runtime/gateway-and-nodes),
[Agent helpers](https://docs.openclaw.ai/plugins/sdk-runtime/agent),
[Sessionextensions](https://docs.openclaw.ai/plugins/hooks/prompt-and-session#session-extensions-and-next-turn-injections),
[Background work](https://docs.openclaw.ai/plugins/sdk-runtime/background-work),
[Toolplugins](https://docs.openclaw.ai/plugins/tool-plugins).
Dokumentabrufreceipt liegt im separaten Evidence-Verzeichnis unter
`concept-docs/receipt.json`; Markdown-Abruf benötigte einen Browser-User-Agent.
SDK-Typen bestätigen Flächen, nicht die erfolgreiche Integration des Plugins.

## 8. Migration und Freigabetests

Legacy-OAuth und bestehende Policy sind gemeinsam explizit einem verifizierten
Profil zuzuordnen. Kein First-Visitor-Claim; keine stillschweigende Team-/Worker-
Freigabe. Bestehende Ressourcenrechte unverändert übernehmen, neue Freigaben
standardmäßig aus. Atomar/idempotent migrieren, alte Zustände gesichert halten,
Konflikte ablehnen. Rollback darf die Verbindung nicht wieder global nutzbar machen.

| Szenario | Erwartetes Ergebnis im empfohlenen Vertrag |
| --- | --- |
| A und B, getrennte Sessions, parallel | Jeweils eigene Bindung, keine Cache-/Tokenvermischung |
| B steuert A in freigegebener gemeinsamer Session | A bleibt aktives Konto; B nutzt nur freigegebene Agentfähigkeiten |
| B will auf eigenes gemeinschaftlich freigegebenes Konto wechseln | Expliziter geprüfter Wechsel zwischen Aufträgen; alte Handles ungültig |
| B kennt fremde Account-ID ohne Freigabe | Ablehnung vor Credentials/Graph |
| Nicht freigegebenes Konto in shared/lesbarer Ausgabegrenze | Keine behauptete Privatheit; sichere Ablehnung bzw. persönlicher Kontext |
| Job auf `graph-a`, keine Live-Person | A nur durch aktuelle Worker-/Autonomiefreigabe, Ersteller nicht geraten |
| Prompt sagt „verwende B“, Worker ist A | Keine Änderung der Bindung; fremdes Ziel kann keine Autorität erzeugen |
| Worker A mit widersprechender Sessionbindung B / UI-Wechsel auf B | `graph_execution_account_conflict` bzw. Wechselablehnung vor Credentials/Graph |
| Delegation an anderen freigegebenen Agent | Dessen eigene Konto-/Dienstrechte; keine automatische Rechteerweiterung |
| Generischer Spawn/Fork ohne nachgewiesene Bindung | Kein Graph-Fallback; expliziter Workerpfad oder Ablehnung |
| Disconnect/Widerruf/Kontowechsel nach Approval | Kein neuer Zugriff mit altem Kontext; genehmigte Parameter nicht umdeuten |
| Manipulierter Sessionextensionwert | Keine Bindung aus Projektion übernehmen |
| Crash/Retry/Parallelität bei Migration oder OAuth-Login | Exklusive generationstreue Veröffentlichung oder sichere Ablehnung |

Nächster Umsetzungsschritt: profilgebundener Credential-/Policy-Layer, zentrale
Ausführungsauflösung, bereinigte UI-Anzeige/Freigabe und statische Workerbindung.
Gezielte Zwei-Benutzer-Hosttests mit synthetischen Graph-Verbindungen prüfen diese
Tabelle vor realer OAuth-Migration und Veröffentlichung. Zusätzliche dynamische
Bindungsvererbung ist keine Voraussetzung für diesen ausdrücklich begrenzten Pfad.

## 9. Status und Verantwortlichkeit

Konzeptowner: Plugin Engineering; Security trägt Kontrollen/Review bei.
Creative Problem Solver lieferte die Problemtrennung; Capability Synthesis die
Kombination aus sessionfester Auswahl und kleinen kontogebundenen Workern.
Kein Skill/Workshop-Produkt erforderlich; Artefakt ist Pluginprojektdokumentation.

Triageentscheidung: **adapt** — OpenClaws sessionfestes Modellkonto-Muster auf
Graph übertragen, aber ohne Modell-Failover und mit expliziten Team-/Workerrechten.
Nichtstun lässt nur den alten Single-Account-Zustand; automatische Live-Sprecher-
umschaltung ist für den allgemeinen Pfad verworfen. Kleine feste Worker sind die
Fallback-Ergänzung, dynamische Bindungsvererbung bleibt zurückgestellt. Nächster
Nachweis gehört Plugin Engineering: zentrale Kontoauflösung und die Testmatrix.

Unabhängiger Konzeptreview am 12.10.2026: eine materielle Präzisierung der
Worker-/Sessionkonfliktregel wurde übernommen und gezielt nachgeprüft.
Terminales Review: Konzept akzeptiert, keine verbleibenden materiellen
Konzeptbefunde; kein Implementierungs- oder Produktions-PASS.

Dieses Konzept grenzt den früheren Feasibility-Blocker ein: Sein Race-Beleg bleibt
für automatische Sprecherumschaltung relevant, blockiert aber nicht ausdrücklich
freigegebene sessionfeste Teamnutzung. Private Live-Requester-Semantik wird nicht
als gelöst dargestellt. Konzeptreview ersetzt keine Implementierungsprüfung.

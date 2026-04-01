# StatikManager V1 — Projektregeln

## Allgemeine Arbeitsregeln

### Reihenfolge: Erst analysieren, dann ändern
Vor jeder Änderung wird der betroffene Code vollständig gelesen und verstanden.
Kein blindes Umschreiben. Kein Raten. Immer zuerst den Ist-Zustand erfassen.

### /original bleibt unberührt
Der Ordner `/original` enthält den unveränderten Quellstand des Projekts.
Keine Datei in `/original` wird jemals verändert, verschoben oder gelöscht.
Er dient als Referenz und Rückfallposition.

### Änderungen nur außerhalb von /original
Alle Anpassungen, Refactorings, neuen Dateien und Migrationen landen ausschließlich
im Arbeitsbereich außerhalb von `/original` (z. B. `/src`, `/migration`, etc.).

### Backup-Prinzip vor größeren Änderungen
Vor strukturell bedeutsamen Änderungen wird ein Snapshot oder Zwischenstand gesichert.
Idealerweise als Kopie im `/migration`-Ordner mit Zeitstempel-Suffix.

### Klare Rollentrennung
| Rolle        | Aufgabe                                      |
|--------------|----------------------------------------------|
| Orchestrator | Koordiniert Aufgaben, delegiert an Agenten  |
| Planner      | Analysiert und zerlegt Aufgaben              |
| Fiona        | Setzt Änderungen um                          |
| Nolen        | Prüft Ergebnisse, erkennt Regressionen       |

Kein Agent überschreitet seine Rolle. Planer ändern keinen Code. Umsetzer planen nicht.

---

## Agentenstruktur

Alle Agentendefinitionen liegen unter `/agents/`.
Skills liegen unter `/skills/`.
Prompts und Vorlagen liegen unter `/prompts/`.
Migrationsschritte werden in `/migration/` dokumentiert.

---

## Kommunikationsformat zwischen Agenten

Jeder Auftrag zwischen Agenten enthält:
- **Ziel**: Was soll erreicht werden?
- **Kontext**: Welche Dateien / Abhängigkeiten sind relevant?
- **Einschränkungen**: Was darf nicht verändert werden?
- **Erfolgskriterium**: Woran erkennt man, dass die Aufgabe erledigt ist?

---

## Git-Regeln

Nach jeder erfolgreichen Änderung IMMER automatisch:

1. git add .
2. git commit -m "Update: kurze Beschreibung der Änderung"
3. git push

WICHTIG:
- Keine Änderung ohne Commit + Push abschließen
- Vor Commit immer Build prüfen
- Keine Nachfrage notwendig

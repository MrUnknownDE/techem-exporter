# Techem Smart Exporter

Ein automatisiertes Tool zum Export von Verbrauchsdaten (Heizung & Warmwasser) aus dem Techem Mieterportal in eine strukturierte Excel-Tabelle. Das Skript nutzt Playwright zur Browser-Automatisierung und navigiert effizient durch die monatlichen Verbrauchsübersichten.

## Features
- **Automatischer Login**: Handhabt den Azure AD B2C Login-Prozess.
- **Cookie-Banner Bypass**: Erkennt und schließt den Cookiebot-Banner automatisch.
- **Intelligente Extraktion**: Nutzt Regex-Parsing, um Heizungs- und Wasserwerte direkt aus dem DOM zu lesen.
- **Smart Formatting**: Erstellt eine Excel-Datei im Horizontal-Layout (Monat | Warmwasser | Heizung).
- **Admin-Ready**: Vollständig über `.env` konfigurierbar mit anpassbaren Timeouts und Debug-Modus.

## Voraussetzungen
- Node.js (v16 oder höher)
- npm (Node Package Manager)

## Nutzung

Starte den Export mit:
```bash
npm i 
node index.js
```

Das Skript öffnet (im Debug-Modus) den Browser, loggt sich ein und iteriert durch alle Monate. Die fertige Datei wird als `Techem_Smart_Export_DATUM.xlsx` im Projektordner gespeichert.

## Troubleshooting

- **Hängt beim Login?**: Setze `DEBUG_MODE=true` und beobachte, ob Techem ein Captcha oder eine MFA-Abfrage verlangt.
- **Keine Daten gefunden?**: Das Skript nutzt Regex zur Suche nach "Heating" und "Water". Sollte Techem die Sprache des Portals ändern, müssen die Begriffe in der `index.js` ggf. angepasst werden.
- **Timeouts?**: Erhöhe `TIMEOUT_PAGE` in der `.env`, falls deine Internetverbindung oder das Techem-Portal langsam reagieren.

## Lizenz
MIT - Feel free to PR :P
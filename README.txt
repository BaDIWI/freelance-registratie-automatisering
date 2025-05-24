# Tentoo Automatisering – Medewerkersregistratie via Python & Selenium

Dit project automatiseert het aanmaken van jobsheets voor freelancers in het Tentoo-platform. Het leest een Excelbestand in met medewerkergegevens, navigeert via Selenium door het platform, vult de juiste gegevens in en geeft live feedback via de terminal. De status van elke invoer wordt direct opgeslagen in het Excelbestand met kleurcodes.

Gebouwd voor werksituaties waarin dagelijks tientallen medewerkers ingevoerd moesten worden — met minimale handmatige handelingen.

---

## 🎯 Functionaliteiten

- Leest medewerkerdata uit `Actieve_Dimona.xlsx` (Medewerker, Plaats, Uren)
- Automatisch invullen van het jobsheet-formulier op my.tentoo.be
- Kleurt Excelcellen automatisch:
  - ✅ Groen = contract succesvol aangemaakt
  - ⚠️ Geel = al ingepland
  - ❌ Rood = fout / medewerker geblokkeerd
- Logging in aparte `.txt` file voor naslag
- Live status in de command line tijdens het proces

---

## 📂 Bestanden

| Bestand | Omschrijving |
|---------|--------------|
| `main event.py` | Hoofdscript met volledige automatisering |
| `Actieve_Dimona.xlsx` | Inputbestand met medewerkers |
| `Namen_Medewerkers.txt` | Logbestand met statusupdates |
| `chromedriver.exe` | Vereist door Selenium voor browserautomatisering |

---

## 🧠 Gebouwd voor de praktijk

Deze tool is ontwikkeld om administratief werk rondom freelancers te versnellen.  
Hij combineert webinteractie, foutcontrole, logging en visuele terugkoppeling in één script dat door een niet-programmeur gestart kan worden.

---

## 🔒 Licentie

**Alle rechten voorbehouden.**  
Dit project is bedoeld als voorbeeld en mag niet worden gekopieerd, aangepast of commercieel gebruikt zonder uitdrukkelijke toestemming van de maker.


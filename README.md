# BET PRO

**Aplicatia:** https://betpro-laamkrsjryitxjuadqeyn2.streamlit.app/

Bot de analiza pentru pariuri pe fotbal. In fiecare dimineata GitHub ruleaza botul singur:

1. verifica pariurile de ieri (rezultat, profit, cota de la start = CLV);
2. ia toate meciurile zilei si toate cotele din API-Football (sute de ligi, cu U19, feminin, ligile 2-3);
3. calculeaza sansele (model Poisson, ca in Excel) si le amesteca 50/50 cu sansa corecta a caselor (Pinnacle sau media caselor);
4. trimite emailul cu: VALUE (3-15%, un pariu pe meci), suspecte, cele mai sigure, bilete pe cote tinta, bombe.

Jurnalul sta in `data/picks.csv` si `data/tickets.csv`, rapoartele in `reports/` (`latest.html` = azi).

## Setari (Settings > Secrets and variables > Actions)

| Secret | Ce pui |
|---|---|
| `API_FOOTBALL_KEY` | cheia API-Football |
| `GMAIL_USER` | adresa Gmail de pe care pleaca emailul |
| `GMAIL_APP_PASSWORD` | parola de aplicatie Gmail (16 litere) |
| `MAIL_TO` | cui se trimite (mai multe adrese cu virgula) |

Rulare manuala: tab-ul **Actions** > **BET PRO zilnic** > **Run workflow**.

`excel/` = versiunea Excel (v35).

Nicio selectie nu e sigura. Joaca responsabil, doar cat iti permiti sa pierzi.

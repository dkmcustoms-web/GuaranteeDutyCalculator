# 🛃 Import Duty Calculator

Streamlit app om douanerechten en BTW te berekenen op basis van commodity code, factuurwaarde en wisselkoers.

## Functies

- **Dropdown met zoekfunctie** – selecteer commodity via code + omschrijving
- **Automatische wisselkoersen** – via InforEuro (officiële EC-koersen) met fallback naar Frankfurter/ECB
- **Manuele koers mogelijk** – overschrijf de live koers indien gewenst
- **Meerdere lijnen** – voeg onbeperkt commodity-regels toe
- **Berekeningen per lijn**:
  - Factuurwaarde × Wisselkoers = Waarde in EUR
  - Waarde EUR × Duty% = Duty Calculated
  - (Waarde EUR + Duty) × 21% = BTW
  - Duty + BTW = Totaal Taxes
- **Totalen** in de header
- **Export** naar Excel (.xlsx) en CSV

## Installatie

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Commodities aanpassen

Bewerk `commodities.csv` met twee kolommen:
| Kolom | Beschrijving |
|-------|-------------|
| `commodity_code` | TARIC/GN-code |
| `description` | Omschrijving |
| `duty_pct` | Invoerrecht in % (bijv. `12.8`) |

## Deployment op Streamlit Cloud

1. Push deze map naar een GitHub repository
2. Ga naar [share.streamlit.io](https://share.streamlit.io)
3. Koppel je GitHub repo
4. Stel in: `Main file path = app.py`
5. Deploy!

## Wisselkoersbronnen

- **Primair**: [InforEuro API](https://ec.europa.eu/budg/inforeuro/api/public/) – officiële maandelijkse EC-boekhoudkoersen
- **Fallback**: [Frankfurter API](https://www.frankfurter.app/) – dagelijkse ECB-referentiekoersen (gratis, geen API-key nodig)

## Bestandsstructuur

```
duty_calculator/
├── app.py            # Hoofdapplicatie
├── style.css         # CSS-opmaak
├── commodities.csv   # Commodity codes en tarieven
├── requirements.txt  # Python-afhankelijkheden
└── README.md         # Deze documentatie
```

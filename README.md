# Henry Hub vs OTE + Forward

Skript `henry_hub_ote_forward.py` vytváří interaktivní HTML report se sekcemi Henry Hub, OTE, Porovnání a **4. Forward**. Data CME načte při vytváření reportu a vloží je přímo do HTML. Pro aktualizaci dat stačí vytvořit nový report stejným příkazem.

Pro spuštění ve Windows:

1. Nainstaluj [Python](https://www.python.org/downloads/) ve verzi 3.10 nebo novější, pokud jej ještě nemáš.
2. Na stránce tohoto repozitáře klikni na zelené **Code → Download ZIP**. Stažený ZIP rozbal.
3. Do rozbalené složky vedle skriptu vlož své původní soubory `Spot plynu - HH.xlsx`, `spotové ceny OTE - plyn.xlsx` a `USD to EUR.xlsx`. Tyto Excely nejsou součástí repozitáře.
4. Otevři tuto složku v Průzkumníku Windows, do horního adresního řádku napiš `cmd` a stiskni Enter. Otevře se příkazový řádek v dané složce.
5. Zadej `py -m pip install -r requirements.txt` a počkej na dokončení instalace. Tento krok stačí při prvním spuštění.
6. Zadej `py henry_hub_ote_forward.py`. Po dokončení se vypíše cesta k vytvořenému reportu a skript skončí.
7. Dvojklikem otevři soubor `henry_hub_ote_interaktivni_analyza.html` v prohlížeči. Python už nemusí běžet.

Graf Forward používá [CME Natural Gas Settlements](https://www.cmegroup.com/markets/energy/natural-gas/natural-gas.settlements.html): **Month** na ose X a **Settle v USD/MMBtu** na ose Y. Kontrakty jsou seřazené podle měsíců dodání. U grafu je datum vypořádání a čas načtení dat. Settle jsou denní vypořádací ceny; jejich dostupnost závisí na publikaci CME.

Pokud se data CME při vytváření reportu nepodaří načíst, sekce Forward zobrazí chybu. Ostatní sekce reportu lze nadále používat. Po obnovení přístupu k CME spusť skript znovu.

Zpracování odpovědí CME bylo ověřeno na vzorových datech. Živý přístup k CME při přípravě nebylo možné ověřit kvůli omezení sítě. Podrobnosti a volitelné parametry jsou v [NAVOD.txt](NAVOD.txt).

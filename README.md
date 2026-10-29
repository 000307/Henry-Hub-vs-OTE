# Henry Hub vs OTE + Forward

Skript `henry_hub_ote_forward.py` vytváří interaktivní HTML report se sekcemi Henry Hub, OTE, Porovnání a **4. Forward**. Data CME a kurzy Fed načte při vytváření reportu a vloží je přímo do HTML. Pro aktualizaci dat stačí vytvořit nový report stejným příkazem.

Pro spuštění ve Windows:

1. Nainstaluj [Python](https://www.python.org/downloads/) ve verzi 3.10 nebo novější, pokud jej ještě nemáš.
2. Na stránce tohoto repozitáře klikni na zelené **Code → Download ZIP**. Stažený ZIP rozbal.
3. Do rozbalené složky vedle skriptu vlož své dva původní soubory `Spot plynu - HH.xlsx` a `spotové ceny OTE - plyn.xlsx`. Tyto Excely nejsou součástí repozitáře.
4. Otevři tuto složku v Průzkumníku Windows, do horního adresního řádku napiš `cmd` a stiskni Enter. Otevře se příkazový řádek v dané složce.
5. Zadej `py -m pip install -r requirements.txt` a počkej na dokončení instalace. Tento krok stačí při prvním spuštění.
6. S připojením k internetu zadej `py henry_hub_ote_forward.py`. Po dokončení se vypíše cesta k vytvořenému reportu a skript skončí.
7. Dvojklikem otevři soubor `henry_hub_ote_interaktivni_analyza.html` v prohlížeči. Python už nemusí běžet.

Společné přepínače **USD/EUR** a **MMBtu/MWh** platí pro všechny čtyři sekce. Převod ceny na EUR používá kurz z [historické tabulky Fed](https://www.federalreserve.gov/RELEASES/H10/hist/dat00_eu.htm): `Rate` je počet USD za 1 EUR, tedy `EUR = USD / Rate`. Pro převod ceny z jednotky MMBtu na MWh se násobí koeficientem `3,4121416`.

Původní publikované číselné kurzy se zachovají. Chybějící dny, včetně `ND` a víkendů, dostanou aritmetický průměr nejbližšího předchozího a následujícího původního kurzu. Celá mezera mezi nimi používá stejnou hodnotu. Pro cenová data před 3. 1. 2000 se použije kurz ze 3. 1. 2000; po posledním publikovaném kurzu poslední dostupný kurz.

Graf Forward používá [CME Natural Gas Settlements](https://www.cmegroup.com/markets/energy/natural-gas/natural-gas.settlements.html): **Month** na ose X a **Settle** na ose Y ve zvolené měně a jednotce. Kontrakty jsou seřazené podle měsíců dodání. Celá křivka v EUR používá kurz k datu vypořádání CME; u grafu je uvedeno datum použitého kurzu. Settle jsou denní vypořádací ceny a jejich dostupnost závisí na publikaci CME.

Pokud se data CME při vytváření reportu nepodaří načíst, sekce Forward zobrazí chybu. Ostatní sekce reportu lze nadále používat. Po obnovení přístupu k CME spusť skript znovu.

Kvůli omezení sítě nebylo v tomto cloudovém prostředí možné živě ověřit přímé stahování z Fed ani CME. Formát primární tabulky Fed byl ověřen přes webový nástroj. Podrobnosti a volitelné parametry jsou v [NAVOD.txt](NAVOD.txt).

Ověřovací testy parseru a doplnění kurzů lze spustit příkazem `py -m unittest discover -s tests -v`.

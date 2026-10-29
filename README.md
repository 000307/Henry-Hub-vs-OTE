# Henry Hub vs OTE + Forward

Skript `henry_hub_ote_forward.py` vytváří interaktivní HTML report se sekcemi Henry Hub, OTE, Porovnání a **4. Forward**. Spotové ceny Henry Hub z EIA, data CME a kurzy Fed načte při vytváření reportu a vloží je přímo do HTML. Pro aktualizaci dat stačí vytvořit nový report stejným příkazem.

Pro spuštění ve Windows:

1. Nainstaluj [Python](https://www.python.org/downloads/) ve verzi 3.10 nebo novější, pokud jej ještě nemáš.
2. Na stránce tohoto repozitáře klikni na zelené **Code → Download ZIP**. Stažený ZIP rozbal.
3. Do rozbalené složky vedle skriptu vlož svůj soubor `spotové ceny OTE - plyn.xlsx`. Je to jediný potřebný místní Excel a není součástí repozitáře.
4. Otevři tuto složku v Průzkumníku Windows, do horního adresního řádku napiš `cmd` a stiskni Enter. Otevře se příkazový řádek v dané složce.
5. Zadej `py -m pip install -r requirements.txt` a počkej na dokončení instalace. **Po stažení této nové verze instalaci zopakuj i v případě, že už jsi knihovny instaloval.** Přibyla knihovna `xlrd>=2.0.1` pro načtení formátu `.xls`.
6. S připojením k internetu zadej `py henry_hub_ote_forward.py`. Po dokončení se vypíše cesta k vytvořenému reportu a skript skončí.
7. Dvojklikem otevři soubor `henry_hub_ote_interaktivni_analyza.html` v prohlížeči. Python už nemusí běžet.

Spotové ceny Henry Hub se stáhnou z [EIA](https://www.eia.gov/dnav/ng/hist/rngwhhdd.htm) přes odkaz [Download Data (.xls)](https://www.eia.gov/dnav/ng/hist_xls/RNGWHHDd.xls). Skript čte list `Data 1` a sloupce `Date` a `Henry Hub Natural Gas Spot Price (Dollars per Million Btu)`.

Společné přepínače **USD/EUR** a **MMBtu/MWh** platí pro všechny čtyři sekce. Převod ceny na EUR používá kurz z [historické tabulky Fed](https://www.federalreserve.gov/RELEASES/H10/hist/dat00_eu.htm): `Rate` je počet USD za 1 EUR, tedy `EUR = USD / Rate`. Pro převod ceny z jednotky MMBtu na MWh se násobí koeficientem `3,4121416`.

Původní publikované číselné kurzy se zachovají. Chybějící dny, včetně `ND` a víkendů, dostanou aritmetický průměr nejbližšího předchozího a následujícího původního kurzu. Celá mezera mezi nimi používá stejnou hodnotu. Pro cenová data před 3. 1. 2000 se použije kurz ze 3. 1. 2000; po posledním publikovaném kurzu poslední dostupný kurz.

Graf Forward používá [CME Natural Gas Settlements](https://www.cmegroup.com/markets/energy/natural-gas/natural-gas.settlements.html): **Month** na ose X a **Settle** na ose Y ve zvolené měně a jednotce. Kontrakty jsou seřazené podle měsíců dodání. Celá křivka v EUR používá kurz k datu vypořádání CME; u grafu je uvedeno datum použitého kurzu. Settle jsou denní vypořádací ceny a jejich dostupnost závisí na publikaci CME.

Pokud se data CME při vytváření reportu nepodaří načíst, sekce Forward zobrazí chybu. Ostatní sekce reportu lze nadále používat. Po obnovení přístupu k CME spusť skript znovu.

Kvůli omezení sítě nebylo v tomto cloudovém prostředí možné živě ověřit stažení binárního souboru EIA ani přímé stahování z Fed a CME. Cílový odkaz EIA a formát primární tabulky Fed byly ověřeny přes webový nástroj. Podrobnosti a volitelné parametry jsou v [NAVOD.txt](NAVOD.txt).

Ověřovací testy načítání cvičných `.xls` sešitů, parseru a doplnění kurzů lze spustit příkazem `py -m unittest discover -s tests -v`.

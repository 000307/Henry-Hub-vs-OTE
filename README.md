# Henry Hub vs OTE + Forward

Skript `henry_hub_ote_forward.py` vytváří interaktivní HTML report se sekcemi Henry Hub, OTE, Porovnání a **4. Forward**. Při vytváření reportu načte spotové ceny Henry Hub z EIA, doplní místní Excel OTE o nově zveřejněné dny a získá data CME i kurzy Fed. Výsledná data vloží přímo do HTML. Pro aktualizaci dat stačí vytvořit nový report stejným příkazem.

Pro spuštění ve Windows:

1. Nainstaluj [Python](https://www.python.org/downloads/) ve verzi 3.10 nebo novější, pokud jej ještě nemáš.
2. Na stránce tohoto repozitáře klikni na zelené **Code → Download ZIP**. Stažený ZIP rozbal.
3. Do rozbalené složky vedle skriptu vlož svůj původní soubor `spotové ceny OTE - plyn.xlsx`. Soubor není součástí repozitáře. Musí obsahovat sloupec `Date` a alespoň jeden původní cenový sloupec OTE v USD/MWh nebo USD/MMBtu.
   Pro přidání forwardových cen OTE do sekce 4 vlož vedle skriptu také `Forwardy OTE.xlsx`. Tento soubor se pouze čte a tvoje přiložená data se do veřejného repozitáře nenahrávají.
4. Otevři tuto složku v Průzkumníku Windows, do horního adresního řádku napiš `cmd` a stiskni Enter. Otevře se příkazový řádek v dané složce.
5. Zadej `py -m pip install -r requirements.txt` a počkej na dokončení instalace. **Po stažení této nové verze instalaci zopakuj i v případě, že už jsi knihovny instaloval.** Knihovna `xlrd>=2.0.1` slouží k načtení formátu EIA `.xls`.
6. S připojením k internetu zadej `py henry_hub_ote_forward.py`. Po dokončení se vypíše cesta k vytvořenému reportu a skript skončí.
7. Dvojklikem otevři soubor `henry_hub_ote_interaktivni_analyza.html` v prohlížeči. Python už nemusí běžet.

Spotové ceny Henry Hub se stáhnou z [EIA](https://www.eia.gov/dnav/ng/hist/rngwhhdd.htm) přes odkaz [Download Data (.xls)](https://www.eia.gov/dnav/ng/hist_xls/RNGWHHDd.xls). Skript čte list `Data 1` a sloupce `Date` a `Henry Hub Natural Gas Spot Price (Dollars per Million Btu)`.

OTE se doplňuje z [ročních zpráv OTE](https://www.ote-cr.cz/cs/statistika/rocni-zprava?date=2026-01-01). Skript podle dnešního data v Praze projde potřebné roky od konce platných cen v tvém Excelu, najde plynový ZIP a z listu `Odchylky NC BAL` načte novější skutečně zveřejněné dny. Cenu v Kč/MWh převede na USD pomocí oficiálních kurzů ČNB; pro víkendy a svátky použije poslední předchozí dostupný kurz.

Pro aktuální rok vybírá skript denně aktualizovanou verzi **V0**. Verze V1 a V2 mají měsíční či čtvrtletní zpoždění. [OTE uvádí dostupnost V0 do D−3](https://www.ote-cr.cz/cs/statistika/rocni-zprava), takže poslední zveřejněná cena může být několik dní před dneškem. U staršího roku lze při nedostupné V0 použít dostupnou vyšší verzi.

**Nové řádky se uloží přímo do původního Excelu OTE.** Před zápisem vznikne vedle souboru samostatná záloha s jedinečným názvem. Historické ceny a ostatní listy zůstanou zachovány. Když nejsou nové ceny, nic se nezapisuje, záloha se nevytváří a kurzy ČNB se nestahují. Pokud aktuální roční zpráva zatím není vydaná, skript to oznámí a použije dostupná data. Neočekávaná chyba při doplnění zastaví program a původní Excel se nepoškodí. Program v tomto případě nepokračuje k vytvoření nového reportu.

Společné přepínače **USD/EUR** a **MMBtu/MWh** platí pro všechny čtyři sekce. Převod ceny na EUR používá kurz z [historické tabulky Fed](https://www.federalreserve.gov/RELEASES/H10/hist/dat00_eu.htm): `Rate` je počet USD za 1 EUR, tedy `EUR = USD / Rate`. Pro převod ceny z jednotky MMBtu na MWh se násobí koeficientem `3,4121416`.

Původní publikované číselné kurzy se zachovají. Chybějící dny, včetně `ND` a víkendů, dostanou aritmetický průměr nejbližšího předchozího a následujícího původního kurzu. Celá mezera mezi nimi používá stejnou hodnotu. Pro cenová data před 3. 1. 2000 se použije kurz ze 3. 1. 2000; po posledním publikovaném kurzu poslední dostupný kurz.

Sekce **4. Forward** porovnává v jednom grafu [CME Natural Gas Settlements](https://www.cmegroup.com/markets/energy/natural-gas/natural-gas.settlements.html) a forwardové ceny OTE z prvního listu `Forwardy OTE.xlsx`. CME používá **Month** na ose X a **Settle** na ose Y. U OTE skript vybere jediný řádek, jehož `Datum` je nejblíže dnešku v Praze; při stejné vzdálenosti přednostně vybere minulý den. Všechny OTE produkty použije z tohoto řádku, chybějící ceny nedoplňuje z jiného dne.

Měsíční OTE produkty se zobrazí jako body, kvartály jako vodorovné úseky přes tři měsíce a produkty `Cal` přes celý rok. Základní jednotka OTE je **EUR/MWh**; obě řady reagují na společné přepínače měny a jednotky. Převod OTE do USD používá kurz Fed k datu vybraného řádku (`USD = EUR × Rate`), zatímco převod CME do EUR používá datum vypořádání CME.

Pokud `Forwardy OTE.xlsx` chybí, report se přesto vytvoří a sekce 4 oznámí chybějící soubor. Jinou cestu lze zadat přes `--ote-forward-input`. Settle CME jsou denní vypořádací ceny a jejich dostupnost závisí na publikaci CME.

Pokud se data CME při vytváření reportu nepodaří načíst, sekce Forward zobrazí chybu. Ostatní sekce reportu lze nadále používat. Po obnovení přístupu k CME spusť skript znovu.

Kvůli omezení sítě nebylo v tomto cloudovém prostředí možné živě ověřit stažení souborů EIA a OTE ani přímé stahování z Fed a CME. Cílový odkaz EIA, primární tabulka Fed a textový formát kurzů ČNB byly ověřeny přes webový nástroj. Podrobnosti a volitelné parametry jsou v [NAVOD.txt](NAVOD.txt).

Ověřovací testy načítání cvičných `.xls`, ZIP a `.xlsx` souborů, doplnění spotových cen OTE, čtení forwardových produktů a převodů kurzů lze spustit příkazem `py -m unittest discover -s tests -v`.

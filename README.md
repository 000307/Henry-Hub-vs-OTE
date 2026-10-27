# Henry Hub vs OTE + Forward

Skript `henry_hub_ote_forward.py` vytváří interaktivní report s původními sekcemi Henry Hub, OTE a Porovnání a novou sekcí **4. Forward**. Report se zobrazí v prohlížeči po spuštění skriptu s tvými vstupními Excely.

Pro spuštění ve Windows:

1. Nainstaluj [Python](https://www.python.org/downloads/) ve verzi 3.10 nebo novější, pokud jej ještě nemáš.
2. Na stránce tohoto repozitáře klikni na zelené **Code → Download ZIP**. Stažený ZIP rozbal.
3. Do rozbalené složky vedle skriptu vlož své původní soubory `Spot plynu - HH.xlsx`, `spotové ceny OTE - plyn.xlsx` a `USD to EUR.xlsx`. Tyto Excely nejsou součástí repozitáře.
4. Otevři tuto složku v Průzkumníku Windows, do horního adresního řádku napiš `cmd` a stiskni Enter. Otevře se příkazový řádek v dané složce.
5. Zadej `py -m pip install -r requirements.txt` a počkej na dokončení instalace. Tento krok stačí při prvním spuštění.
6. Zadej `py henry_hub_ote_forward.py`. Report se otevře v prohlížeči. Pokud se neotevře, zkopíruj do prohlížeče celou adresu `http://127.0.0.1:…`, kterou skript vypíše.

**Příkazový řádek s Pythonem nech otevřený.** Při otevřeném reportu se data CME kontrolují každých 5 minut. Ukončení skriptu: Ctrl+C. Samostatné otevření vygenerovaného HTML dvojklikem zobrazí uložený snímek; automatické aktualizace vyžadují běžící Python a vypsanou lokální adresu.

Graf Forward používá [CME Natural Gas Settlements](https://www.cmegroup.com/markets/energy/natural-gas/natural-gas.settlements.html): **Month** na ose X a **Settle v USD/MMBtu** na ose Y. Kontrakty jsou seřazené podle měsíců dodání. Settle jsou denní vypořádací ceny; nová data se objeví po publikaci CME. Při výpadku zůstanou poslední úspěšně načtená data a zobrazí se upozornění.

Zpracování odpovědí, cache a aktualizace v prohlížeči byly ověřeny na vzorových datech. Živý přístup k CME zde kvůli omezení sítě nebylo možné ověřit. Podrobnosti a volitelné parametry jsou v `NAVOD.txt`.

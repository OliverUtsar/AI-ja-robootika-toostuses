#!/usr/bin/env python3
"""
Akna detailide arvutamine tootmisreeglite põhjal koos kliendi erisoovide toega.
"""

import pandas as pd
import re

def loe_tootmisreeglid(fail: str = 'Tootmisreeglid.xlsx') -> pd.DataFrame:
    """Loeb tootmisreeglid Exceli failist."""
    return pd.read_excel(fail)

def arvuta_valem(valem: str, laius: int, korgus: int, kogus: int = 1, tkx_korrutis: int = None, tkx_ylal_korrutis: int = None) -> str:
    """Arvutab valemi põhjal väärtuse."""
    if pd.isna(valem) or valem == 'tk':
        return None
    
    # Asenda L ja K vastavate väärtustega
    valem = valem.replace('L', str(laius)).replace('K', str(korgus))
    
    # Arvuta valemi tulemus
    try:
        tulemus = int(eval(valem))
    except:
        return None
    
    # Lisa sulgudesse korrutis, kui on tkx ja tkx_ylal_korrutis on olemas
    if tkx_korrutis is not None and tkx_ylal_korrutis is not None:
        korrutis = tkx_korrutis * tkx_ylal_korrutis * kogus
        return f"{tulemus} ({korrutis}tk)"
    elif tkx_korrutis is not None:
        korrutis = tkx_korrutis * kogus
        return f"{tulemus} ({korrutis}tk)"
    else:
        return tulemus

def arvuta_kohandatud_klaasiliistud(vertikaalsed: int, horisontaalsed: int, laius: int, korgus: int) -> str:
    """Arvutab kohandatud klaasiliistude konfiguratsiooni."""
    if vertikaalsed == 0 and horisontaalsed == 0:
        return 'Puudub (kliendi soov)'
    
    # Arvuta vertikaalsed liistud
    vertikaalsed_moot = []
    if vertikaalsed > 0:
        vahe = (laius - 80) // (vertikaalsed + 1)  # 80mm on raami paksus
        for i in range(1, vertikaalsed + 1):
            positsioon = 40 + i * vahe  # 40mm on raami servast
            vertikaalsed_moot.append(f"VERT {positsioon}")
    
    # Arvuta horisontaalsed liistud
    horisontaalsed_moot = []
    if horisontaalsed > 0:
        vahe = (korgus - 96) // (horisontaalsed + 1)  # 96mm on raami paksus
        for i in range(1, horisontaalsed + 1):
            positsioon = 48 + i * vahe  # 48mm on raami servast
            horisontaalsed_moot.append(f"HOR {positsioon}")
    
    kogu_liistud = vertikaalsed_moot + horisontaalsed_moot
    if not kogu_liistud:
        return 'Puudub (kliendi soov)'
    
    return ', '.join(kogu_liistud)

def arvuta_detailid(df: pd.DataFrame, toote_tyyp: str, laius: int, korgus: int, kasi: str, kogus: int = 1, erisoovid: dict = None) -> dict:
    """Arvutab akna detailid tootmisreeglite põhjal, arvestades kliendi erisoove."""
    detailid = {}
    
    # Vaikimisi erisoovid
    if erisoovid is None:
        erisoovid = {}
    
    # Käsitsi erisoove
    klaasiliistud_soov = erisoovid.get('klaasiliistud', None)
    vertikaalsed_liistud = erisoovid.get('vertikaalsed_liistud', None)
    horisontaalsed_liistud = erisoovid.get('horisontaalsed_liistud', None)
    
    for i in range(len(df)):
        rida = df.iloc[i]
        
        # Kontrolli, kas see on konfigureerimise rida
        if pd.notna(rida['A/U']) and rida['A/U'] in ['A', 'TA']:
            if rida['toote tyyp'] == toote_tyyp and rida['käsi'] == kasi:
                # Arvuta põhiinfo - alati lisame need väärtused
                detailid['toote_tyyp'] = toote_tyyp
                detailid['laius'] = laius
                detailid['korgus'] = korgus
                detailid['kogus'] = kogus
                detailid['kasi'] = kasi
                
                # Arvuta raami detailid
                if pd.notna(rida['raam']):
                    detailid['raam_laius'] = rida['raam']
                
                if pd.notna(rida['Unnamed: 7']):
                    detailid['raam_korgus'] = rida['Unnamed: 7']
                
                # Arvuta klaasi/KLP mõõdud
                if pd.notna(rida['klaasi/KLP mõõt']):
                    detailid['klaasi_laius'] = rida['klaasi/KLP mõõt']
                
                if pd.notna(rida['Unnamed: 10']):
                    detailid['klaasi_korgus'] = rida['Unnamed: 10']
                
                # Arvuta järgmise rea detailid
                if i + 1 < len(df):
                    jargmine_rida = df.iloc[i + 1]
                    
                    # Arvuta raami laius
                    if pd.notna(jargmine_rida['raam']):
                        tkx_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 10']):
                            tkx = str(jargmine_rida['Unnamed: 10'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        detailid['raam_laius_arvutatud'] = arvuta_valem(jargmine_rida['raam'], laius, korgus, kogus, tkx_korrutis)
                        
                    # Arvuta raami kõrgus
                    if pd.notna(jargmine_rida['Unnamed: 7']):
                        tkx_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 11']):
                            tkx = str(jargmine_rida['Unnamed: 11'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        detailid['raam_korgus_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 7'], laius, korgus, kogus, tkx_korrutis)
                        
                    # Arvuta klaasi/KLP mõõdud
                    if pd.notna(jargmine_rida['klaasi/KLP mõõt']):
                        tkx_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 11']):
                            tkx = str(jargmine_rida['Unnamed: 11'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        detailid['klaasi_laius_arvutatud'] = arvuta_valem(jargmine_rida['klaasi/KLP mõõt'], laius, korgus, kogus, tkx_korrutis)
                        
                    # Arvuta klaasi/KLP kõrgus
                    if pd.notna(jargmine_rida['Unnamed: 10']):
                        tkx_korrutis = None
                        if pd.notna(jargmine_rida['Unnamed: 18']):
                            tkx = str(jargmine_rida['Unnamed: 18'])
                            if 'tkx' in tkx:
                                tkx_korrutis = int(re.search(r'tkx(\d+)', tkx).group(1))
                        detailid['klaasi_korgus_arvutatud'] = arvuta_valem(jargmine_rida['Unnamed: 10'], laius, korgus, kogus, tkx_korrutis)
                        
                    # Arvuta klaasiliistud
                    if pd.notna(jargmine_rida['klaasiliistud']):
                        # Rakenda kliendi erisoove
                        if vertikaalsed_liistud is not None or horisontaalsed_liistud is not None:
                            # Kohandatud vertikaalsed/horisontaalsed liistud
                            vert = vertikaalsed_liistud if vertikaalsed_liistud is not None else 0
                            hor = horisontaalsed_liistud if horisontaalsed_liistud is not None else 0
                            detailid['klaasiliistud_arvutatud'] = arvuta_kohandatud_klaasiliistud(vert, hor, laius, korgus)
                        elif klaasiliistud_soov is not None:
                            if klaasiliistud_soov.lower() == 'puudub':
                                detailid['klaasiliistud_arvutatud'] = 'Puudub (kliendi soov)'
                            else:
                                detailid['klaasiliistud_arvutatud'] = klaasiliistud_soov
                        else:
                            # Kui kliendil pole erisoove, arvutame siiski väärtuse valemi põhjal
                            detailid['klaasiliistud_arvutatud'] = arvuta_valem(jargmine_rida['klaasiliistud'], laius, korgus, kogus)
    
    return detailid

def main():
    """Peamine funktsioon kasutajaliidesega."""
    print("Akna detailide arvutamine (koos kliendi erisoovide toega)")
    print("=" * 60)
    
    akna_tyyp = input("Sisesta akna tüüp (tavaline/topelt): ").lower()
    
    # Kontrolli toote tüüpi kohe pärast sisestamist
    while True:
        toote_tyyp = input("Sisesta toote tüüp (nt 40STAND60x63, 77STAND60x70): ")
        df = loe_tootmisreeglid()
        valid_tooted = df['toote tyyp'].dropna().unique()
        
        if toote_tyyp in valid_tooted:
            break
        else:
            print(f"VIGA: Toode '{toote_tyyp}' ei ole kehtiv!")
            print("Palun sisestage toote tüüp formaadis NNNSTANDXXxXX, näiteks:")
            print("  - 40STAND60x63")
            print("  - 77STAND60x70")
            print("\nSaadaolevad toote tüübid:")
            for tyyp in valid_tooted:
                print(f"  - {tyyp}")
            print()
    
    laius = int(input("Sisesta laius (mm): "))
    korgus = int(input("Sisesta kõrgus (mm): "))
    kogus = int(input("Sisesta akende kogus: "))
    kasi = input("Sisesta käelisus (P/V): ").upper()
    
    # Küsime kliendi erisoove
    print("\nKliendi erisoovid klaasiliistude kohta:")
    print("-" * 60)
    
    # Kõigepealt küsime, kas kliend üldse soovib klaasiliiste
    klaasiliistud_soov = input("Kas soovite klaasiliiste? (jah/ei): ").strip().lower()
    
    # Koosta erisoovide sõnastik
    erisoovid = {}
    
    if klaasiliistud_soov == 'jah':
        # Kui jah, küsime täpsemalt
        print("\nMäärake klaasiliistude konfiguratsioon:")
        erisoov_vertikaalsed = input("Vertikaalsed klaasiliistud (arv): ").strip()
        erisoov_horisontaalsed = input("Horisontaalsed klaasiliistud (arv): ").strip()
        
        if erisoov_vertikaalsed:
            erisoovid['vertikaalsed_liistud'] = int(erisoov_vertikaalsed)
        if erisoov_horisontaalsed:
            erisoovid['horisontaalsed_liistud'] = int(erisoov_horisontaalsed)
    elif klaasiliistud_soov == 'ei':
        # Kui ei, märgime, et klaasiliiste ei soovita
        erisoovid['klaasiliistud'] = 'puudub'
    # Kui kasutaja jätab tühjaks, kasutame vaikimisi väärtusi
    
    df = loe_tootmisreeglid()
    detailid = arvuta_detailid(df, toote_tyyp, laius, korgus, kasi, kogus, erisoovid)
    
    print("\n" + "=" * 60)
    print("AKNA DETAILID (ARVESTADES KLIENDI ERISOOVE)")
    print("=" * 60)
    print(f"Akende kogus: {kogus}")
    print(f"Kliendi erisoovid: {erisoovid if erisoovid else 'Puuduvad'}")
    print()
    
    # Grupeerime mõõdud loogilistesse kategooriatesse
    print("1. PÕHIMÕÕDUD (ANDMEBAASIST - STANDARDSED VÄÄRTUSED)")
    print("-" * 60)
    if 'toote_tyyp' in detailid:
        print(f"Toote tüüp: {detailid['toote_tyyp']}")
    if 'laius' in detailid:
        print(f"Akna laius: {detailid['laius']} mm")
    if 'korgus' in detailid:
        print(f"Akna kõrgus: {detailid['korgus']} mm")
    if 'kasi' in detailid:
        print(f"Käelisus: {detailid['kasi']}")
    if 'raam_laius' in detailid:
        print(f"Raami laius (standard): {detailid['raam_laius']} mm")
    if 'raam_korgus' in detailid:
        print(f"Raami kõrgus (standard): {detailid['raam_korgus']} mm")
    if 'klaasi_laius' in detailid:
        print(f"Klaasi laius (standard): {detailid['klaasi_laius']} mm")
    if 'klaasi_korgus' in detailid:
        print(f"Klaasi kõrgus (standard): {detailid['klaasi_korgus']} mm")
    
    print()
    print("2. ARVUTATUD MÕÕDUD (KLIENDI SISESTATUD MÕÕTUDE PÕHAL)")
    print("-" * 60)
    if 'raam_laius_arvutatud' in detailid:
        print(f"Raami laius (arvutatud): {detailid['raam_laius_arvutatud']}")
    if 'raam_korgus_arvutatud' in detailid:
        print(f"Raami kõrgus (arvutatud): {detailid['raam_korgus_arvutatud']}")
    if 'klaasi_laius_arvutatud' in detailid:
        print(f"Klaasi laius (arvutatud): {detailid['klaasi_laius_arvutatud']}")
    if 'klaasi_korgus_arvutatud' in detailid:
        print(f"Klaasi kõrgus (arvutatud): {detailid['klaasi_korgus_arvutatud']}")
    
    print()
    print("3. KLAASILIISTUD (KLIENDI ERISOOVID)")
    print("-" * 60)
    if 'klaasiliistud_arvutatud' in detailid:
        if detailid['klaasiliistud_arvutatud'] == 'Puudub (kliendi soov)':
            print("Klaasiliistud: Kliendil pole soovi klaasiliistude järele")
        else:
            print(f"Klaasiliistud: {detailid['klaasiliistud_arvutatud']}")
            if isinstance(detailid['klaasiliistud_arvutatud'], str) and 'VERT' in detailid['klaasiliistud_arvutatud']:
                print("  (V = vertikaalne liist, H = horisontaalne liist)")
    
    print("=" * 60)

if __name__ == '__main__':
    main()
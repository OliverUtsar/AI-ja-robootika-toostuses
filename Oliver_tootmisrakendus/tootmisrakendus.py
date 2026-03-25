#!/usr/bin/env python3
import os
import pandas as pd
from typing import Dict, List, Optional, Union


# Esimesed 6 veergu on toote põhiandmed
TOOTE_VEERUD = [
    "A/U",
    "toote tyyp",
    "laius",
    "kõrgus",
    "tk tyyptellimusel",
    "käsi",
]


class Tootmisrakendus:
    def __init__(self, fail: str = "Tootmisreeglid.xlsx"):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        fail_path = os.path.join(script_dir, fail)

        # Loeme kahepäiselise tabelina (1. ja 2. rida on päised)
        self.df = pd.read_excel(fail_path, header=[0, 1])

        # Puhastame päised
        self.df.columns = pd.MultiIndex.from_tuples(
            [
                (
                    str(l0).strip() if pd.notna(l0) else "",
                    str(l1).strip() if pd.notna(l1) else "",
                )
                for (l0, l1) in self.df.columns.to_list()
            ]
        )

        self.tooted: Dict[str, List[Dict]] = {}
        self._lae_tooted()

    # -----------------------------
    #  VALEMITE ARVUTAMINE
    # -----------------------------
    def _eval_valem(self, expr: str, L: float, K: float, tk: int):
        """Arvutab valemi stringist, kasutades ainult L, K ja tk."""
        if not isinstance(expr, str):
            return expr

        s = expr.replace("L", str(L)).replace("K", str(K))
        s = s.replace("tkx4", str(tk * 4))
        s = s.replace("tkx2", str(tk * 2))
        s = s.replace("tk", str(tk))

        try:
            return eval(s)
        except:
            return expr

    # -----------------------------
    #  TOODETE LAADIMINE
    # -----------------------------
    def _lae_tooted(self):
        rows = self.df

        # Käime ridade kaupa: numbrid (0), valemid (1), numbrid (2), valemid (3) jne
        for idx in range(0, len(rows), 2):
            num_row = rows.iloc[idx]

            # Kui valemirida puudub → lõpetame
            if idx + 1 >= len(rows):
                break

            valem_row = rows.iloc[idx + 1]

            # Toote põhiandmed tulevad numbrilisest reast
            a_u = num_row.get(("A/U", "Unnamed: 0_level_1"), None)
            toote_tyyp = num_row.get(("toote tyyp", "Unnamed: 1_level_1"), None)

            if pd.isna(a_u) or pd.isna(toote_tyyp):
                continue

            # Need kolm väärtust on valemite alus
            L = float(num_row.get(("laius", "Unnamed: 2_level_1"), 0))
            K = float(num_row.get(("kõrgus", "Unnamed: 3_level_1"), 0))
            tk = int(num_row.get(("tk tyyptellimusel", "Unnamed: 4_level_1"), 0))

            käsi = num_row.get(("käsi", "Unnamed: 5_level_1"), None)

            detailid: Dict[str, Dict] = {}

            # -----------------------------
            #  DETAILIDE DÜNAAMILINE PARSIMINE
            # -----------------------------
            columns = list(self.df.columns)
            i = 0

            while i < len(columns):
                grp, field = columns[i]

                # Jäta toote põhiandmed ja tühjad grupid vahele
                if grp in TOOTE_VEERUD or grp == "":
                    i += 1
                    continue

                # Jäta "konfig." täielikult välja
                if field.lower().startswith("konfig"):
                    i += 1
                    continue

                # Loo grupp kui seda pole
                if grp not in detailid:
                    detailid[grp] = {"mõõdud": {}, "kogus": {}}

                # Kui see on mõõdu veerg (mitte tk)
                if not field.lower().startswith("tk"):
                    valem = valem_row[(grp, field)]
                    if not pd.isna(valem):
                        detailid[grp]["mõõdud"][field] = self._eval_valem(valem, L, K, tk)

                    # Kontrolli, kas järgmine veerg on kogus
                    if i + 1 < len(columns):
                        next_grp, next_field = columns[i + 1]

                        if next_grp == grp and next_field.lower().startswith("tk"):
                            valem_kogus = valem_row[(next_grp, next_field)]
                            if not pd.isna(valem_kogus):
                                detailid[grp]["kogus"][field] = self._eval_valem(
                                    valem_kogus, L, K, tk
                                )

                            i += 2
                            continue

                    i += 1
                    continue

                # Kui on koguse veerg, aga mõõt oli juba töödeldud → edasi
                i += 1

            # Salvestame toote
            toote_entry = {
                "toote_tyyp": toote_tyyp,
                "A/U": a_u,
                "mõõdud": {"laius": L, "kõrgus": K, "tk": tk},
                "detailid": {"käsi": käsi, "grupid": detailid},
            }

            key = str(toote_tyyp)
            self.tooted.setdefault(key, []).append(toote_entry)

    # -----------------------------
    #  TOOTE DETAILIDE PÄRING
    # -----------------------------
    def get_toote_detailid(
        self, toote_tyyp: str, laius: Optional[Union[int, float]] = None
    ) -> Dict:

        if toote_tyyp not in self.tooted:
            raise ValueError(f"Toodet {toote_tyyp} ei leitud.")

        kandidaadid = self.tooted[toote_tyyp]

        if laius is None:
            return kandidaadid[0]

        for t in kandidaadid:
            if t["mõõdud"]["laius"] == float(laius):
                return t

        raise ValueError(f"Toodet {toote_tyyp} laiusega {laius} ei leitud.")

    def get_all_tooted(self) -> List[str]:
        return list(self.tooted.keys())


# -----------------------------
#  JSON VÄLJUND
# -----------------------------
if __name__ == "__main__":
    import json

    rakendus = Tootmisrakendus()

    tooted_json = {}
    for toode in rakendus.get_all_tooted():
        for variant in rakendus.tooted[toode]:
            key = f'{toode}_{int(variant["mõõdud"]["laius"])}x{int(variant["mõõdud"]["kõrgus"])}'
            tooted_json[key] = variant

    with open("tootmisreeglid_output.json", "w", encoding="utf-8") as f:
        json.dump(tooted_json, f, indent=2, ensure_ascii=False)

    print("Tootmisreeglid on salvestatud faili 'tootmisreeglid_output.json'.")

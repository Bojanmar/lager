import re
import numpy as np
import pandas as pd

# ======================================================
# 1) Normalizacija teksta / jedinica
# ======================================================

def norm_text(s):
    if pd.isna(s):
        return ""
    s = str(s).replace('"', ' ').replace("'", ' ')
    return re.sub(r"\s+", " ", s.strip()).lower()


def norm_unit(u):
    u = norm_text(u)
    repl = {
        "m^2": "m2",
        "m²": "m2",
        "m^1": "m",
        "m¹": "m",
        "m1": "m",
        "kom (rolni)": "kom",
        "kom (rolna)": "kom",
        "kom (pak)": "kom",
        "kg.": "kg",
        "l": "lit",
        "litar": "lit",
        "litra": "lit",
        "litara": "lit",
        "džak": "dzak",
        "djak": "dzak",
        "rolni": "rolna",
        "kom rupa": "kom",
    }
    return repl.get(u, u)


def canon_mat(name: str) -> str:
    """Kanonizacija imena materijala preko alias mape."""
    k = norm_text(name)
    return ALIASES.get(k, k)


# ======================================================
# 2) Markup za iste jedinice
# ======================================================

same_unit_markup = {
    "borner multiplex av 4": 0.20,
    "volteco, volgrip h.1.10 light": 0.15,
    "vintex mp fr 1.5mm": 0.10,
    "geotekstil 300gr-m2": 0.10,
    "poliesterska tkanina filc 100%, rolna 2x1m": 0.10,
}

# ======================================================
# 3) Helper pravila – svako ima _rule_desc za pregled
# ======================================================

def rule_factor_per(area_unit, out_unit, factor):
    def _f(q, uf, ul):
        if norm_unit(uf) != area_unit or norm_unit(ul) != out_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * factor, f"factor {factor} {out_unit}/{area_unit}"

    _f._rule_desc = f"factor_per: {factor} {out_unit}/{area_unit}"
    return _f


def rule_factor_per_len(len_unit, out_unit, factor):
    def _f(q, uf, ul):
        if norm_unit(uf) != len_unit or norm_unit(ul) != out_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * factor, f"factor {factor} {out_unit}/{len_unit}"

    _f._rule_desc = f"factor_per_len: {factor} {out_unit}/{len_unit}"
    return _f


def rule_per_piece(out_unit, factor):
    def _f(q, uf, ul):
        if norm_unit(uf) != "kom" or norm_unit(ul) != out_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * factor, f"{factor} {out_unit}/kom"

    _f._rule_desc = f"per_piece: {factor} {out_unit}/kom"
    return _f


def rule_m2_to_rolna(m2_per_rolna, extra=0.0):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m2" or norm_unit(ul) != "rolna":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        rolls = (q / m2_per_rolna) * (1.0 + extra)
        return rolls, f"m2→rolna; {m2_per_rolna} m2/rolna; +{int(extra * 100)}%"

    _f._rule_desc = f"m2_to_rolna: {m2_per_rolna} m2/rolna, +{int(extra*100)}%"
    return _f


def rule_piece_to_kg(kg_per_kom):
    f = rule_per_piece("kg", kg_per_kom)
    f._rule_desc = f"piece_to_kg: {kg_per_kom} kg/kom"
    return f


def rule_identity(from_unit, to_unit):
    def _f(q, uf, ul):
        if norm_unit(uf) != from_unit or norm_unit(ul) != to_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q, f"1:1 {from_unit}->{to_unit}"

    _f._rule_desc = f"identity: {from_unit}->{to_unit}"
    return _f


def rule_m2_to_lit(liters_per_m2):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m2" or norm_unit(ul) != "lit":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * liters_per_m2, f"{liters_per_m2} lit/m2"

    _f._rule_desc = f"m2_to_lit: {liters_per_m2} lit/m2"
    return _f


def rule_m_to_lit(liters_per_m):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m" or norm_unit(ul) != "lit":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * liters_per_m, f"{liters_per_m} lit/m"

    _f._rule_desc = f"m_to_lit: {liters_per_m} lit/m"
    return _f


def rule_m_to_rolna(m_per_rolna, extra=0.0):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m" or norm_unit(ul) != "rolna":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        rolls = (q / m_per_rolna) * (1.0 + extra)
        return rolls, f"m→rolna; {m_per_rolna} m/rolna; +{int(extra * 100)}%"

    _f._rule_desc = f"m_to_rolna: {m_per_rolna} m/rolna, +{int(extra*100)}%"
    return _f


def rule_kg_to_dzak(kg_per_dzak):
    def _f(q, uf, ul):
        if norm_unit(uf) != "kg" or norm_unit(ul) != "dzak":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q / kg_per_dzak, f"1/{kg_per_dzak} dzak/kg"

    _f._rule_desc = f"kg_to_dzak: 1/{kg_per_dzak} dzak/kg"
    return _f


# ======================================================
# 4) Osnovne norme – material_rules + extend_rules
# ======================================================

material_rules = {}


def extend_rules(name, rules_to_add):
    key = norm_text(name)
    material_rules.setdefault(key, []).extend(rules_to_add)


# -- osnovna pravila iz tvog koda --
extend_rules("alchimica aqua smart dur 2k",
             [rule_factor_per("m2", "kg", 0.20)])
extend_rules("hyperdesmo pb 2k a+b, 20+20lit",
             [rule_factor_per("m2", "kg", 3)])
extend_rules("alchimica water foam 1k lv",
             [rule_piece_to_kg(0.3)])
extend_rules("waterfoam catalyst 1 kg",
             [rule_per_piece("kg", 0.015)])
extend_rules("aquasmart - pb 1k 10kg, kg",
             [rule_factor_per("m2", "kg", 1.5),
              rule_factor_per_len("m", "kg", 0.7)])
extend_rules("borner gebortol vs",
             [rule_factor_per("m2", "kg", 0.4)])
extend_rules("cold cure polyurea 2k  a+b",
             [rule_factor_per("m2", "kg", 2.0)])
extend_rules("dual seal 15mil lg 8,92m2 rolna",
             [rule_m2_to_rolna(8.92, extra=0.20)])

extend_rules("hydrobloc 575 integral - 1k pu resin elastic 6.5kg",
             [rule_piece_to_kg(0.3)])
extend_rules("hydrocat 514 - highly active accelerator",
             [rule_piece_to_kg(0.002)])
extend_rules("hydrobloc 510 - second-foam",
             [rule_piece_to_kg(0.3)])

extend_rules("hyperdesmo -  ady-e 4lit",
             [rule_factor_per("m2", "kg", 0.15)])
extend_rules("hyperdesmo grey 1k 25 kg kanta",
             [rule_factor_per("m2", "kg", 2.5)])
extend_rules("hyperseal 2k f, 12kg",
             [rule_factor_per_len("m", "kg", 2.5)])
extend_rules("illbruck pu901 600ml, kom",
             [rule_per_piece("kg", 0.045)])
extend_rules("microsealer pu, 20kg",
             [rule_factor_per("m2", "kg", 0.4),
              rule_factor_per_len("m", "kg", 0.2)])
extend_rules("resin bau creck flex 2k a+b, 10+10.8kg",
             [rule_piece_to_kg(0.3)])
extend_rules("resin bau hydrogum, 20kg",
             [rule_piece_to_kg(0.3)])
extend_rules("resin bau water stopper 20kg comp a, 1,4kg comp b",
             [rule_piece_to_kg(0.3)])
extend_rules("stopaq 2100 aquastop 0,53kg, kom",
             [rule_piece_to_kg(0.003)])

extend_rules("vandex am 10, 20kg džak",
             [rule_factor_per("m3", "kg", 6.0)])
extend_rules("vandex bb 75, 25kg",
             [rule_factor_per("m2", "kg", 3.4)])
extend_rules("vandex cemelast liquid 9kg",
             [rule_factor_per("m2", "kg", 3.4)])
extend_rules("vandex injection mortar (vim) 25kg",
             [rule_piece_to_kg(0.11)])
extend_rules("vandex plug, 15kg kanta",
             [])
extend_rules("vandex uni moratar 1z 25kg",
             [rule_factor_per_len("m", "kg", 5.0),
              rule_piece_to_kg(0.5)])
extend_rules("yapseal 106, komp a, 20kg",
             [rule_factor_per("m2", "kg", 3.5)])
extend_rules("yapseal 106, komp b, 10kg",
             [rule_factor_per("m2", "kg", 1.75)])

# ======================================================
# 5) Proširenja i konstante
# ======================================================

PB2K_DENSITY_KG_PER_L = 1.15
PB2K_LIT_PER_M2 = round(2.4 / PB2K_DENSITY_KG_PER_L, 3)  # ≈ 2.087 L/m2
ME508_3x25m_ROLL_LENGTH_M = 25
ME508_4x50m_ROLL_LENGTH_M = 50
CEMFLEX_PLATE_M_PER_KOM = 2
OMEGA_HOLDERS_PER_PAK = 100
OMEGA_HOLDERS_PER_M = 5
OMEGA_PAK_PER_M = OMEGA_HOLDERS_PER_M / OMEGA_HOLDERS_PER_PAK  # 0.05
VANDEX_PLUG_KG_PER_KOM = 0.1
VOLGRIP_WIDTH_M = 0.10
DUR2K_KG_PER_KOM = 0.20
DUR2K_KG_PER_M = 0.20
DUALSEAL_STD_M_PER_ROLNA = 8.92


ALIASES = {
    norm_text("Auqa Smart DUR 2k, 4+4kg"): norm_text("alchimica aqua smart dur 2k"),
}


# dodatna pravila i override-i
extend_rules("Cleaner, 5kg", [rule_per_piece("kg", 0.025)])

extend_rules("ILLBRUCK PU901 600ml, kom", [
    rule_factor_per_len("m", "kom", 0.07),
    rule_factor_per("m2", "kom", 0.05),
])

extend_rules("Illbruck ME 508 privremena traka 3x25m", [
    rule_factor_per("m2", "m", 1.0),
    rule_factor_per_len("m", "kom", 1.0 / ME508_3x25m_ROLL_LENGTH_M),
    rule_factor_per("m2", "kom", 1.0 / ME508_3x25m_ROLL_LENGTH_M),
])
extend_rules("Illbruck ME 508 privremena traka 4x50m", [
    rule_factor_per("m2", "m", 1.0),
    rule_factor_per_len("m", "kom", 1.0 / ME508_4x50m_ROLL_LENGTH_M),
    rule_factor_per("m2", "kom", 1.0 / ME508_4x50m_ROLL_LENGTH_M),
])

extend_rules("Auqa Smart DUR 2k, 4+4kg", [
    rule_factor_per("m2", "kg", 0.20),
    rule_factor_per_len("m", "kg", DUR2K_KG_PER_M),
    rule_per_piece("kg", DUR2K_KG_PER_KOM),
])
extend_rules("alchimica aqua smart dur 2k", [
    rule_factor_per("m2", "kg", 0.20),
])

extend_rules("HYPERDESMO PB 2K A+B, 20+20lit", [
    rule_m2_to_lit(PB2K_LIT_PER_M2),
    rule_m_to_lit(PB2K_LIT_PER_M2),
])

extend_rules("Alumanation 301 FT, kanta", [rule_identity("kom", "kanta")])

extend_rules("CemFLEX VB Coated Steel Plate  15cm X 2m",
             [rule_per_piece("m", CEMFLEX_PLATE_M_PER_KOM)])
extend_rules("CemFLEX VB Omega holder  100kom pak", [
    rule_per_piece("pak", 1 / OMEGA_HOLDERS_PER_PAK),
    rule_factor_per_len("m", "pak", OMEGA_PAK_PER_M),
])

extend_rules("DUAL SEAL 15mil, STD 8,92m2 rolna", [
    rule_m2_to_rolna(8.92, extra=0.00),
    rule_m_to_rolna(DUALSEAL_STD_M_PER_ROLNA, extra=0.00),
])

extend_rules("HYPERDESMO 2K-W  Comp A+B, 1.5+7.5kg",
             [rule_factor_per("m2", "kg", 1.0)])

extend_rules("HYPERDESMO GREY 1k 25 kg kanta", [
    rule_per_piece("kg", 25.0),
    rule_factor_per_len("m", "kg", 2.5),
])

extend_rules("RESIN BAU Easy Inject, 20kg",
             [rule_per_piece("kg", 0.09)])

extend_rules("CONNECT KSKSEAL privremena traka", [
    rule_factor_per_len("m", "kom", 0.07),
    rule_factor_per("m2", "kom", 0.05),
])

extend_rules("VANDEX Injection Mortar (VIM) 25kg",
             [rule_per_piece("kg", 0.11)])

extend_rules("VANDEX PLUG, 15kg kanta",
             [rule_per_piece("kg", VANDEX_PLUG_KG_PER_KOM)])

extend_rules("VANDEX SUPER, (25kg)",
             [rule_factor_per("m2", "kg", 2.0)])

extend_rules("VOLTECO, Volgrip H.1.10 light",
             [rule_factor_per_len("m", "m2", VOLGRIP_WIDTH_M)])

# DUAL SEAL PARAGRANULAR override
_dual_par_key = norm_text("DUAL SEAL PARAGRANULAR, 23kg džak")
material_rules[_dual_par_key] = [
    rule_per_piece("kg", 23.0),
    rule_kg_to_dzak(23.0),
    rule_identity("kom", "dzak"),
]

# ======================================================
# 6) Injektiranje aktivnih prodora – recepti
# ======================================================

def _n(s):
    return norm_text(s)


INJEKT_TRIGGER = _n("Injektiranje aktivnih prodora")

INJEKT_RECEPT_DEFAULT = {
    _n("ALU Packer 10/100 mm, kom"): (1.0, "kom")
}

INJEKT_RECEPT_BY_RACUN = {
    _n("03-25"): {
        _n("ALU Packer 10/100 mm, kom"): (1.00, "kom"),
        _n("Waterfoam 1K LV, 20 kg"): (70 / 74, "kg"),
        _n("Waterfoam Catalyst 1 kg"): (3.5 / 74, "kg"),
        _n("RESIN BAU HydroGum, 20kg"): (20 / 74, "kg"),
        _n("RESIN BAU Easy Inject, 20kg"): (40 / 74, "kg"),
        _n("STOPAQ 2100 AQUASTOP 0,53KG, kom"): (3 / 74, "kom"),
        _n("VANDEX PLUG, 15kg kanta"): (15 / 74, "kg"),
    },
    _n("04-25"): {
        _n("ALU Packer 10/100 mm, kom"): (1.00, "kom"),
        _n("Waterfoam 1K LV, 20 kg"): (200 / 360, "kg"),
        _n("Waterfoam Catalyst 1 kg"): (10 / 360, "kg"),
        _n("RESIN BAU Easy Inject, 20kg"): (60 / 360, "kg"),
        _n("VANDEX PLUG, 15kg kanta"): (50 / 360, "kg"),
    },
    _n("05-25"): {
        _n("ALU Packer 10/100 mm, kom"): (1.00, "kom"),
        _n("RESIN BAU HydroGum, 20kg"): (40 / 33, "kg"),
    },
    _n("06-25"): {
        _n("ALU Packer 10/100 mm, kom"): (1.00, "kom"),
        _n("RESIN BAU Creck Flex 2k A+B, 10+10.8kg"): (40 / 14, "kg"),
    },
    _n("07-25"): {
        _n("ALU Packer 10/100 mm, kom"): (500 / 509, "kom"),
        _n("Paker Poljska 10x110mm SWG"): (60 / 509, "kom"),
        _n("Paker Poljska 10x300mm SWG"): (80 / 509, "kom"),
        _n("Cleaner, 5kg"): (10 / 509, "kg"),
        _n("Waterfoam 1K LV, 20 kg"): (240 / 509, "kg"),
        _n("Waterfoam Catalyst 1 kg"): (24 / 509, "kg"),
        _n("RESIN BAU Creck Flex 2k A+B, 10+10.8kg"): (320 / 509, "kg"),
        _n("RESIN BAU Easy Inject, 20kg"): (20 / 509, "kg"),
        _n("HydroBloc 575 Integral - 1K PU Resin Elastic 6.5kg"): (19.5 / 509, "kg"),
        _n("VANDEX PLUG, 15kg kanta"): (90 / 509, "kg"),
    },
}


def _broj_pakera_po_setu(df):
    brojevi = {}
    for racun, grp in df.groupby("Broj računa", dropna=False):
        g = grp.copy()
        g["_mat_norm"] = g["Materijal"].map(_n)
        g["_jm_norm"] = g["Jedinica mere za fakturisanje"].map(norm_unit)
        alu_mask = g["_mat_norm"].eq(_n("ALU Packer 10/100 mm, kom"))
        if alu_mask.any():
            base = pd.to_numeric(
                g.loc[alu_mask, "Količina za fakturisanje"],
                errors="coerce"
            ).max()
        else:
            base = pd.to_numeric(
                g.loc[g["_jm_norm"].eq("kom"), "Količina za fakturisanje"],
                errors="coerce"
            ).max()
        brojevi[_n(racun)] = base if pd.notna(base) else None
    return brojevi


def _override_injekt(row, mat_norm, uf_norm, ul_norm, broj_pakera_map):
    tip = row.get("Pozicija za fakturisanje - tip hidroizolacije")
    if pd.isna(tip) or INJEKT_TRIGGER not in _n(tip):
        return None, None
    racun = _n(row.get("Broj računa"))
    base = broj_pakera_map.get(racun)
    if base is None or base == 0:
        return None, None
    rec = INJEKT_RECEPT_BY_RACUN.get(racun, {})
    if mat_norm in rec:
        faktor, out_u = rec[mat_norm]
        return base * faktor, f"recipe_injekt[{row.get('Broj računa')}]: {faktor} {out_u}/packer"
    if mat_norm in INJEKT_RECEPT_DEFAULT:
        faktor, out_u = INJEKT_RECEPT_DEFAULT[mat_norm]
        return base * faktor, f"recipe_injekt_default: {faktor} {out_u}/packer"
    return None, None


# ======================================================
# 7) Override – Sanacija kapilarne vlage
# ======================================================

def _override_sanacija(row, mat, uf, ul):
    tip_hidro = row.get("Pozicija za fakturisanje - tip hidroizolacije")
    if pd.notna(tip_hidro) and norm_text(tip_hidro) == norm_text("Sanacija kapilarne vlage"):
        _uni_alias = {
            norm_text("VANDEX UNI MORTAR 1Z 25kg"),
            norm_text("VANDEX UNI MORATAR 1Z 25kg"),
        }
        if mat in _uni_alias and uf == "m2" and ul == "kg":
            return row["Količina za fakturisanje"] * 4.0, "override(sanacija kapilarne vlage): 4 kg/m2"
    return None, None


# ======================================================
# 8) Kalkulacija količine za skidanje
# ======================================================

pair_defaults = {("m2", "kg"): 1.5, ("m", "kg"): 0.30, ("kom", "kg"): 0.30}


def make_heur_rule(uf, ul):
    uf = norm_unit(uf)
    ul = norm_unit(ul)
    if (uf, ul) not in pair_defaults:
        return None
    f = pair_defaults[(uf, ul)]
    if (uf, ul) == ("m2", "kg"):
        return rule_factor_per("m2", "kg", f)
    if (uf, ul) == ("m", "kg"):
        return rule_factor_per_len("m", "kg", f)
    if (uf, ul) == ("kom", "kg"):
        return rule_per_piece("kg", f)


def calc_skidanje(row, broj_pakera_map):
    mat_raw = row["Materijal"]
    qty = row["Količina za fakturisanje"]
    u_fakt = row["Jedinica mere za fakturisanje"]
    u_lager = row["Jedinica mere za lager - skidanje količine"]
    if pd.isna(mat_raw) or pd.isna(qty) or pd.isna(u_fakt) or pd.isna(u_lager):
        return pd.Series([pd.NA, "nedostaju_podaci"])

    mat = canon_mat(mat_raw)
    uf = norm_unit(u_fakt)
    ul = norm_unit(u_lager)

    out, note = _override_injekt(row, mat, uf, ul, broj_pakera_map)
    if out is not None:
        return pd.Series([out, note])

    out, note = _override_sanacija(row, mat, uf, ul)
    if out is not None:
        return pd.Series([out, note])

    if uf == ul:
        base = qty
        add = same_unit_markup.get(mat, 0.0)
        if add:
            return pd.Series([base * (1 + add), f"same_unit +{int(add * 100)}%"])
        return pd.Series([base, "same_unit"])

    rules = material_rules.get(mat, [])
    if not rules:
        return pd.Series([pd.NA, f"nema_pravila({u_fakt}->{u_lager})"])

    for r in rules:
        out, note = r(qty, uf, ul)
        if out is not None:
            return pd.Series([out, f"rule: {note}"])

    return pd.Series([pd.NA, f"pravilo_ne_pokriva({u_fakt}->{u_lager})"])


def calc_skidanje_all(row, broj_pakera_map, material_rules_all):
    mat_raw = row["Materijal"]
    qty = row["Količina za fakturisanje"]
    u_fakt = row["Jedinica mere za fakturisanje"]
    u_lager = row["Jedinica mere za lager - skidanje količine"]
    if pd.isna(mat_raw) or pd.isna(qty) or pd.isna(u_fakt) or pd.isna(u_lager):
        return pd.Series([pd.NA, "nedostaju_podaci"])

    mat = canon_mat(mat_raw)
    uf = norm_unit(u_fakt)
    ul = norm_unit(u_lager)

    out, note = _override_injekt(row, mat, uf, ul, broj_pakera_map)
    if out is not None:
        return pd.Series([out, note])

    out, note = _override_sanacija(row, mat, uf, ul)
    if out is not None:
        return pd.Series([out, note])

    if uf == ul:
        base = qty
        add = same_unit_markup.get(mat, 0.0)
        if add:
            return pd.Series([base * (1 + add), f"same_unit +{int(add * 100)}%"])
        return pd.Series([base, "same_unit"])

    rules = material_rules_all.get(mat, [])
    if not rules:
        return pd.Series([pd.NA, f"nema_pravila({u_fakt}->{u_lager})"])

    for r in rules:
        out, note = r(qty, uf, ul)
        if out is not None:
            tag = "heur" if (mat in material_rules_all and r not in material_rules.get(mat, [])) else "rule"
            return pd.Series([out, f"{tag}: {note}"])

    return pd.Series([pd.NA, f"pravilo_ne_pokriva({u_fakt}->{u_lager})"])


# ======================================================
# 9) Koeficijenti – analiza uporedbe
# ======================================================

def _n2(s):
    if pd.isna(s):
        return ""
    s = str(s).lower().strip()
    return re.sub(r"\s+", " ", s)


flex_materials = {
    _n2("ALCHIMICA Water Foam 1K"),
    _n2("HydroBloc 575 Integral - 1K PU Resin Elastic 6.5kg"),
    _n2("HydroBloc 510 - Second-foam"),
    _n2("RESIN BAU Creck Flex 2k A+B, 10+10.8kg"),
    _n2("RESIN BAU HydroGum, 20kg"),
    _n2("RESIN BAU Water Stopper 20kg COMP A, 1,4kg comp B"),
}
DEFAULT_MIN, DEFAULT_MAX = 0.5, 2.0
FLEX_MIN, FLEX_MAX = 0.1, 5.0


def _calc_coef(row):
    pot, lag = row["Ukupna_potrošnja"], row["Stanje_na_lageru"]
    mat_norm = _n2(row["Materijal"])
    is_flex = mat_norm in flex_materials
    lo, hi = (FLEX_MIN, FLEX_MAX) if is_flex else (DEFAULT_MIN, DEFAULT_MAX)
    if pd.isna(pot) or pot == 0 or pd.isna(lag):
        return pd.Series([np.nan, lo, hi, is_flex, "nema_baze_za_coef"])
    raw = lag / pot
    if raw < lo or raw > hi:
        return pd.Series([raw, lo, hi, is_flex, f"EKSTREMNO raw={raw:.4f} (van {lo}–{hi})"])
    return pd.Series([raw, lo, hi, is_flex, f"ok raw={raw:.4f}"])


# ======================================================
# 10) Glavna funkcija – procesiraj_obracun
# ======================================================

def procesiraj_obracun(lager_file, fakture_file):
    """
    Glavna funkcija koju koristi Streamlit app.
    Prima dva upload-ovana fajla (BytesIO) i vraća:
      - uporedba (koeficijenti_i_uporedba)
      - ekstremni (ekstremni materijali)
      - df_fakture_posle (fakture sa kolonom 'Količina za skidanje sa lagera')
    """
    # 0) Učitavanje Excel fajlova
    df_uvoz = pd.read_excel(lager_file)
    df_fakture = pd.read_excel(fakture_file)

    # 1) Mapiranje JM iz df_uvoz u df_fakture
    mapa_jedinica = df_uvoz.iloc[0].to_dict()
    mapa_materijali = {col: mapa_jedinica[col] for col in df_uvoz.columns}
    df_fakture["Jedinica mere za lager - skidanje količine"] = (
        df_fakture["Materijal"].map(mapa_materijali)
    )

    # 2) Override problematičnih JM
    mask_dual = df_fakture["Materijal"].map(norm_text).eq(_dual_par_key)
    df_fakture.loc[mask_dual, "Jedinica mere za lager - skidanje količine"] = "dzak"

    me508_3 = norm_text("Illbruck ME 508 privremena traka 3x25m")
    me508_4 = norm_text("Illbruck ME 508 privremena traka 4x50m")
    mask_me3 = df_fakture["Materijal"].map(norm_text).eq(me508_3)
    mask_me4 = df_fakture["Materijal"].map(norm_text).eq(me508_4)
    df_fakture.loc[mask_me3 | mask_me4, "Jedinica mere za lager - skidanje količine"] = "kom"

    # 3) Injekt – mapiranje broja pakera po računu
    broj_pakera_map = _broj_pakera_po_setu(df_fakture)

    # 4) Prvi prolaz
    df_fakture_pre = df_fakture.copy()
    df_fakture_pre[["Količina za skidanje sa lagera",
                    "Napomena konverzije"]] = df_fakture_pre.apply(
        lambda row: calc_skidanje(row, broj_pakera_map),
        axis=1
    )
    neuspeh_pre = df_fakture_pre[df_fakture_pre["Količina za skidanje sa lagera"].isna()].copy()

    # 5) Heuristike
    material_rules_all = {k: v[:] for k, v in material_rules.items()}
    for _, r in neuspeh_pre[[
        "Materijal",
        "Jedinica mere za fakturisanje",
        "Jedinica mere za lager - skidanje količine"
    ]].drop_duplicates().iterrows():
        rr = make_heur_rule(r["Jedinica mere za fakturisanje"],
                            r["Jedinica mere za lager - skidanje količine"])
        if rr is not None:
            material_rules_all.setdefault(canon_mat(r["Materijal"]), []).append(rr)

    # 6) Drugi prolaz
    df_fakture_posle = df_fakture.copy()
    df_fakture_posle[["Količina za skidanje sa lagera",
                      "Napomena konverzije"]] = df_fakture_posle.apply(
        lambda row: calc_skidanje_all(row, broj_pakera_map, material_rules_all),
        axis=1
    )

    # 7) PB2K spec (m2/m -> lit)
    PB2K_NAME_NORM = norm_text("HYPERDESMO PB 2K A+B, 20+20lit")
    materijal_norm = df_fakture_posle["Materijal"].map(norm_text)
    jm_fakt_norm = df_fakture_posle["Jedinica mere za fakturisanje"].map(norm_unit)
    jm_lager_norm = df_fakture_posle["Jedinica mere za lager - skidanje količine"].map(norm_unit)

    mask_pb2k = materijal_norm.eq(PB2K_NAME_NORM)
    mask_pb2k_m2 = mask_pb2k & jm_fakt_norm.eq("m2") & jm_lager_norm.eq("lit")
    df_fakture_posle.loc[mask_pb2k_m2, "Količina za skidanje sa lagera"] = (
        pd.to_numeric(
            df_fakture_posle.loc[mask_pb2k_m2, "Količina za fakturisanje"],
            errors="coerce"
        ) * PB2K_LIT_PER_M2
    )
    mask_pb2k_m = mask_pb2k & jm_fakt_norm.eq("m") & jm_lager_norm.eq("lit")
    df_fakture_posle.loc[mask_pb2k_m, "Količina za skidanje sa lagera"] = (
        pd.to_numeric(
            df_fakture_posle.loc[mask_pb2k_m, "Količina za fakturisanje"],
            errors="coerce"
        ) * PB2K_LIT_PER_M2
    )

    # Za 'kom' zaokruži
    mask_kom_final = jm_lager_norm.eq("kom")
    df_fakture_posle.loc[mask_kom_final, "Količina za skidanje sa lagera"] = pd.to_numeric(
        df_fakture_posle.loc[mask_kom_final, "Količina za skidanje sa lagera"],
        errors="coerce"
    ).round(0)

    # 8) SUME i uporedba
    df_fakture_posle["Količina za skidanje sa lagera"] = pd.to_numeric(
        df_fakture_posle["Količina za skidanje sa lagera"], errors="coerce"
    )

    potrosnja = (
        df_fakture_posle.groupby("Materijal", dropna=False)["Količina za skidanje sa lagera"]
        .sum(min_count=1)
        .reset_index()
        .rename(columns={"Količina za skidanje sa lagera": "Ukupna_potrošnja"})
    )

    lager_stanje = df_uvoz.iloc[1].to_frame().reset_index()
    lager_stanje.columns = ["Materijal", "Stanje_na_lageru"]
    lager_stanje["Stanje_na_lageru"] = pd.to_numeric(
        lager_stanje["Stanje_na_lageru"], errors="coerce"
    )

    uporedba = pd.merge(potrosnja, lager_stanje, on="Materijal", how="outer")
    uporedba["Razlika_pre"] = uporedba["Stanje_na_lageru"] - uporedba["Ukupna_potrošnja"]

    # 9) Koeficijenti
    uporedba[["Koeficijent", "Coef_min", "Coef_max", "Flex", "Napomena_coef"]] = uporedba.apply(
        _calc_coef, axis=1
    )
    ekstremni = uporedba[
        uporedba["Napomena_coef"].str.contains("EKSTREMNO", na=False)
    ].copy()

    return uporedba, ekstremni, df_fakture_posle


# ======================================================
# 11) Pregled pravila – za Streamlit sidebar
# ======================================================

def get_rules_overview():
    """
    Vrati DataFrame sa listom svih material_rules i kratkim opisom.
    Ovo koristimo samo za pregled u sidebar-u.
    """
    rows = []
    for mat, rules in material_rules.items():
        for idx, r in enumerate(rules, start=1):
            desc = getattr(r, "_rule_desc", r.__name__)
            rows.append({
                "Materijal (norm)": mat,
                "Rule #": idx,
                "Opis pravila": desc,
            })
    return pd.DataFrame(rows)

import io

def export_to_excel(uporedba, ekstremni, df_fakture_posle):
    """ Priprema fajla za download u Streamlit-u """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        uporedba.to_excel(writer, index=False, sheet_name="koeficijenti_i_uporedba")
        ekstremni.to_excel(writer, index=False, sheet_name="ekstremni")
        df_fakture_posle.to_excel(writer, index=False, sheet_name="fakture_obracun")
    buffer.seek(0)
    return buffer
import streamlit as st

def rules_editor_ui():
    global material_rules  # radimo sa globalnim pravilima

    # pretvorimo pravila u tabelu koja može da se edituje
    data = []
    for mat, rules in material_rules.items():
        for r in rules:
            desc = getattr(r, "_rule_desc", "manual_rule")
            data.append({
                "Materijal": mat,
                "Opis": desc
            })
    df = pd.DataFrame(data)

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        num_rows="dynamic",
        height=300
    )

    if st.button("💾 Sačuvaj izmene"):
        # Ovde čuvamo izmene -> kasnije možemo i import/export JSON
        st.session_state["material_rules_edits"] = edited_df
        st.success("✔ Pravila su sačuvana i biće primenjena na sledećem obračunu.")

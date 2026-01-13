import re
import io
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
        "m^2": "m2", "m²": "m2",
        "m^1": "m",  "m¹": "m", "m1": "m",
        "kom (rolni)": "kom", "kom (rolna)": "kom", "kom (pak)": "kom",
        "kg.": "kg",
        "l": "lit", "litar": "lit", "litra": "lit", "litara": "lit",
        "džak": "dzak", "djak": "dzak",
        "rolni": "rolna",
        "kom rupa": "kom",
    }
    return repl.get(u, u)


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
# 3) Helper pravila – svaki rule nosi _meta + _rule_desc
# ======================================================

def _attach_meta(fn, rule_type, from_unit, to_unit, factor=None, extra=None, enabled=True):
    fn._meta = {
        "rule_type": rule_type,
        "from_unit": from_unit,
        "to_unit": to_unit,
        "factor": factor,
        "extra": extra,
        "enabled": enabled
    }
    if rule_type in ("factor_per", "factor_per_len"):
        fn._rule_desc = f"{rule_type}: {factor} {to_unit}/{from_unit}"
    elif rule_type == "per_piece":
        fn._rule_desc = f"per_piece: {factor} {to_unit}/kom"
    elif rule_type == "identity":
        fn._rule_desc = f"identity: {from_unit}->{to_unit}"
    elif rule_type == "m2_to_rolna":
        fn._rule_desc = f"m2_to_rolna: {factor} m2/rolna, +{int((extra or 0)*100)}%"
    elif rule_type == "m_to_rolna":
        fn._rule_desc = f"m_to_rolna: {factor} m/rolna, +{int((extra or 0)*100)}%"
    elif rule_type == "m2_to_lit":
        fn._rule_desc = f"m2_to_lit: {factor} lit/m2"
    elif rule_type == "m_to_lit":
        fn._rule_desc = f"m_to_lit: {factor} lit/m"
    elif rule_type == "kg_to_dzak":
        fn._rule_desc = f"kg_to_dzak: 1/{factor} dzak/kg"
    else:
        fn._rule_desc = rule_type
    return fn


def rule_factor_per(area_unit, out_unit, factor):
    def _f(q, uf, ul):
        if norm_unit(uf) != area_unit or norm_unit(ul) != out_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * factor, f"factor {factor} {out_unit}/{area_unit}"
    return _attach_meta(_f, "factor_per", area_unit, out_unit, factor=factor)


def rule_factor_per_len(len_unit, out_unit, factor):
    def _f(q, uf, ul):
        if norm_unit(uf) != len_unit or norm_unit(ul) != out_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * factor, f"factor {factor} {out_unit}/{len_unit}"
    return _attach_meta(_f, "factor_per_len", len_unit, out_unit, factor=factor)


def rule_per_piece(out_unit, factor):
    def _f(q, uf, ul):
        if norm_unit(uf) != "kom" or norm_unit(ul) != out_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * factor, f"{factor} {out_unit}/kom"
    return _attach_meta(_f, "per_piece", "kom", out_unit, factor=factor)


def rule_identity(from_unit, to_unit):
    def _f(q, uf, ul):
        if norm_unit(uf) != from_unit or norm_unit(ul) != to_unit:
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q, f"1:1 {from_unit}->{to_unit}"
    return _attach_meta(_f, "identity", from_unit, to_unit, factor=None)


def rule_m2_to_rolna(m2_per_rolna, extra=0.0):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m2" or norm_unit(ul) != "rolna":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        rolls = (q / m2_per_rolna) * (1.0 + extra)
        return rolls, f"m2→rolna; {m2_per_rolna} m2/rolna; +{int(extra * 100)}%"
    return _attach_meta(_f, "m2_to_rolna", "m2", "rolna", factor=m2_per_rolna, extra=extra)


def rule_m_to_rolna(m_per_rolna, extra=0.0):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m" or norm_unit(ul) != "rolna":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        rolls = (q / m_per_rolna) * (1.0 + extra)
        return rolls, f"m→rolna; {m_per_rolna} m/rolna; +{int(extra * 100)}%"
    return _attach_meta(_f, "m_to_rolna", "m", "rolna", factor=m_per_rolna, extra=extra)


def rule_m2_to_lit(liters_per_m2):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m2" or norm_unit(ul) != "lit":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * liters_per_m2, f"{liters_per_m2} lit/m2"
    return _attach_meta(_f, "m2_to_lit", "m2", "lit", factor=liters_per_m2)


def rule_m_to_lit(liters_per_m):
    def _f(q, uf, ul):
        if norm_unit(uf) != "m" or norm_unit(ul) != "lit":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q * liters_per_m, f"{liters_per_m} lit/m"
    return _attach_meta(_f, "m_to_lit", "m", "lit", factor=liters_per_m)


def rule_kg_to_dzak(kg_per_dzak):
    def _f(q, uf, ul):
        if norm_unit(uf) != "kg" or norm_unit(ul) != "dzak":
            return None, f"pravilo_ne_važi({uf}->{ul})"
        return q / kg_per_dzak, f"1/{kg_per_dzak} dzak/kg"
    return _attach_meta(_f, "kg_to_dzak", "kg", "dzak", factor=kg_per_dzak)


# ======================================================
# 4) Pravila: material_rules + alias
# ======================================================

material_rules = {}

def extend_rules(name, rules_to_add):
    key = norm_text(name)
    material_rules.setdefault(key, []).extend(rules_to_add)


ALIASES = {
    norm_text("Auqa Smart DUR 2k, 4+4kg"): norm_text("alchimica aqua smart dur 2k"),
}

def canon_mat(name: str) -> str:
    k = norm_text(name)
    return ALIASES.get(k, k)


# --- osnovna pravila ---
extend_rules("alchimica aqua smart dur 2k", [rule_factor_per("m2", "kg", 0.20)])
extend_rules("hyperdesmo pb 2k a+b, 20+20lit", [rule_factor_per("m2", "kg", 3)])
extend_rules("alchimica water foam 1k lv", [rule_per_piece("kg", 0.3)])
extend_rules("waterfoam catalyst 1 kg", [rule_per_piece("kg", 0.015)])
extend_rules("aquasmart - pb 1k 10kg, kg", [rule_factor_per("m2", "kg", 1.5), rule_factor_per_len("m", "kg", 0.7)])
extend_rules("borner gebortol vs", [rule_factor_per("m2", "kg", 0.4)])
extend_rules("cold cure polyurea 2k  a+b", [rule_factor_per("m2", "kg", 2.0)])
extend_rules("dual seal 15mil lg 8,92m2 rolna", [rule_m2_to_rolna(8.92, extra=0.20)])

extend_rules("hydrobloc 575 integral - 1k pu resin elastic 6.5kg", [rule_per_piece("kg", 0.3)])
extend_rules("hydrocat 514 - highly active accelerator", [rule_per_piece("kg", 0.002)])
extend_rules("hydrobloc 510 - second-foam", [rule_per_piece("kg", 0.3)])

extend_rules("hyperdesmo -  ady-e 4lit", [rule_factor_per("m2", "kg", 0.15)])
extend_rules("hyperdesmo grey 1k 25 kg kanta", [rule_factor_per("m2", "kg", 2.5)])
extend_rules("hyperseal 2k f, 12kg", [rule_factor_per_len("m", "kg", 2.5)])
extend_rules("illbruck pu901 600ml, kom", [rule_per_piece("kg", 0.045)])
extend_rules("microsealer pu, 20kg", [rule_factor_per("m2", "kg", 0.4), rule_factor_per_len("m", "kg", 0.2)])
extend_rules("resin bau creck flex 2k a+b, 10+10.8kg", [rule_per_piece("kg", 0.3)])
extend_rules("resin bau hydrogum, 20kg", [rule_per_piece("kg", 0.3)])
extend_rules("resin bau water stopper 20kg comp a, 1,4kg comp b", [rule_per_piece("kg", 0.3)])
extend_rules("stopaq 2100 aquastop 0,53kg, kom", [rule_per_piece("kg", 0.003)])

extend_rules("vandex am 10, 20kg džak", [rule_factor_per("m3", "kg", 6.0)])
extend_rules("vandex bb 75, 25kg", [rule_factor_per("m2", "kg", 3.4)])
extend_rules("vandex cemelast liquid 9kg", [rule_factor_per("m2", "kg", 3.4)])
extend_rules("vandex injection mortar (vim) 25kg", [rule_per_piece("kg", 0.11)])
extend_rules("vandex plug, 15kg kanta", [])
extend_rules("vandex uni moratar 1z 25kg", [rule_factor_per_len("m", "kg", 5.0), rule_per_piece("kg", 0.5)])
extend_rules("yapseal 106, komp a, 20kg", [rule_factor_per("m2", "kg", 3.5)])
extend_rules("yapseal 106, komp b, 10kg", [rule_factor_per("m2", "kg", 1.75)])


# ======================================================
# 5) Proširenja i konstante
# ======================================================

PB2K_DENSITY_KG_PER_L = 1.15
PB2K_LIT_PER_M2 = round(2.4 / PB2K_DENSITY_KG_PER_L, 3)

ME508_3x25m_ROLL_LENGTH_M = 25
ME508_4x50m_ROLL_LENGTH_M = 50
CEMFLEX_PLATE_M_PER_KOM = 2
OMEGA_HOLDERS_PER_PAK = 100
OMEGA_HOLDERS_PER_M = 5
OMEGA_PAK_PER_M = OMEGA_HOLDERS_PER_M / OMEGA_HOLDERS_PER_PAK
VANDEX_PLUG_KG_PER_KOM = 0.1
VOLGRIP_WIDTH_M = 0.10
DUR2K_KG_PER_KOM = 0.20
DUR2K_KG_PER_M = 0.20
DUALSEAL_STD_M_PER_ROLNA = 8.92

# dodatna pravila
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
# 6) Injektiranje – recepti
# ======================================================

def _n(s): return norm_text(s)

INJEKT_TRIGGER = _n("Injektiranje aktivnih prodora")
INJEKT_RECEPT_DEFAULT = {_n("ALU Packer 10/100 mm, kom"): (1.0, "kom")}

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
            base = pd.to_numeric(g.loc[alu_mask, "Količina za fakturisanje"], errors="coerce").max()
        else:
            base = pd.to_numeric(g.loc[g["_jm_norm"].eq("kom"), "Količina za fakturisanje"], errors="coerce").max()
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
    return None


def calc_skidanje(row, broj_pakera_map, rules_dict):
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

    rules = rules_dict.get(mat, [])
    if not rules:
        return pd.Series([pd.NA, f"nema_pravila({u_fakt}->{u_lager})"])

    for r in rules:
        meta = getattr(r, "_meta", {})
        if meta and meta.get("enabled") is False:
            continue
        out, note = r(qty, uf, ul)
        if out is not None:
            return pd.Series([out, f"rule: {note}"])

    return pd.Series([pd.NA, f"pravilo_ne_pokriva({u_fakt}->{u_lager})"])


# ======================================================
# 9) Koeficijenti – analiza uporedbe (tvoj postojeći)
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
# 10) RULES <-> DF (za editor)
# ======================================================

RULE_TYPES = [
    "factor_per",
    "factor_per_len",
    "per_piece",
    "identity",
    "m2_to_rolna",
    "m_to_rolna",
    "m2_to_lit",
    "m_to_lit",
    "kg_to_dzak",
]

def rules_to_df(rules_dict):
    rows = []
    for mat, rules in rules_dict.items():
        for r in rules:
            meta = getattr(r, "_meta", None)
            if not meta:
                rows.append({
                    "Materijal": mat,
                    "rule_type": "factor_per",
                    "from_unit": "",
                    "to_unit": "",
                    "factor": None,
                    "extra": 0.0,
                    "enabled": True,
                })
                continue
            rows.append({
                "Materijal": mat,
                "rule_type": meta.get("rule_type", ""),
                "from_unit": meta.get("from_unit", ""),
                "to_unit": meta.get("to_unit", ""),
                "factor": meta.get("factor", None),
                "extra": meta.get("extra", 0.0),
                "enabled": meta.get("enabled", True),
            })
    return pd.DataFrame(rows)


def make_rule_from_row(row):
    rt = row["rule_type"]
    fu = row.get("from_unit", "")
    tu = row.get("to_unit", "")
    factor = row.get("factor", None)
    extra = row.get("extra", 0.0)
    enabled = bool(row.get("enabled", True))

    if rt == "factor_per":
        fn = rule_factor_per(fu, tu, float(factor))
    elif rt == "factor_per_len":
        fn = rule_factor_per_len(fu, tu, float(factor))
    elif rt == "per_piece":
        fn = rule_per_piece(tu, float(factor))
    elif rt == "identity":
        fn = rule_identity(fu, tu)
    elif rt == "m2_to_rolna":
        fn = rule_m2_to_rolna(float(factor), extra=float(extra))
    elif rt == "m_to_rolna":
        fn = rule_m_to_rolna(float(factor), extra=float(extra))
    elif rt == "m2_to_lit":
        fn = rule_m2_to_lit(float(factor))
    elif rt == "m_to_lit":
        fn = rule_m_to_lit(float(factor))
    elif rt == "kg_to_dzak":
        fn = rule_kg_to_dzak(float(factor))
    else:
        return None

    fn._meta["enabled"] = enabled
    return fn


def apply_rules_df(base_rules_dict, edited_df):
    if edited_df is None or edited_df.empty:
        return {k: v[:] for k, v in base_rules_dict.items()}

    df = edited_df.copy()
    for c in ["Materijal", "rule_type", "from_unit", "to_unit"]:
        df[c] = df[c].fillna("").astype(str)

    df["enabled"] = df["enabled"].fillna(True).astype(bool)
    df["extra"] = pd.to_numeric(df["extra"], errors="coerce").fillna(0.0)
    if "factor" in df.columns:
        df["factor"] = pd.to_numeric(df["factor"], errors="coerce")

    out = {}
    for mat, grp in df.groupby("Materijal", dropna=False):
        rules = []
        for _, r in grp.iterrows():
            fn = make_rule_from_row(r.to_dict())
            if fn is not None:
                rules.append(fn)
        out[mat] = rules

    for mat, rules in base_rules_dict.items():
        if mat not in out:
            out[mat] = rules[:]
    return out


# ======================================================
# 11) POPIS PARSER (stvarno stanje magacina)
# ======================================================

def parse_stvarno_stanje_excel(stvarno_file, value_row_idx=1):
    """
    Očekivan format:
      - kolone = materijali
      - red 0 = jedinica
      - red 1 = količina (stvarno stanje)

    Vraća df: [Materijal, Stanje_na_magacinu]
    """
    df = pd.read_excel(stvarno_file)
    stanje = df.iloc[value_row_idx].to_frame().reset_index()
    stanje.columns = ["Materijal", "Stanje_na_magacinu"]
    stanje["Stanje_na_magacinu"] = pd.to_numeric(stanje["Stanje_na_magacinu"], errors="coerce")
    return stanje


# ======================================================
# 12) Glavna funkcija – procesiraj_obracun (+ POPIS)
# ======================================================

def procesiraj_obracun(lager_file, fakture_file, edited_rules_df=None, stvarno_file=None):
    """
    Prima:
      - lager_file (ULAZ)
      - fakture_file (IZLAZ)
      - edited_rules_df (opciono)
      - stvarno_file (POPIS / stvarno stanje magacina - opciono)

    Vraća:
      - uporedba
      - ekstremni
      - df_fakture_posle
      - rules_used_df
    """
    df_uvoz = pd.read_excel(lager_file)
    df_fakture = pd.read_excel(fakture_file)

    # primeni editovana pravila (ako postoje)
    base_rules = material_rules
    rules_dict = apply_rules_df(base_rules, edited_rules_df)

    # 1) mapiranje JM iz lagera
    mapa_jedinica = df_uvoz.iloc[0].to_dict()
    df_fakture["Jedinica mere za lager - skidanje količine"] = df_fakture["Materijal"].map(mapa_jedinica)

    # 2) override JM
    mask_dual = df_fakture["Materijal"].map(norm_text).eq(_dual_par_key)
    df_fakture.loc[mask_dual, "Jedinica mere za lager - skidanje količine"] = "dzak"

    me508_3 = norm_text("Illbruck ME 508 privremena traka 3x25m")
    me508_4 = norm_text("Illbruck ME 508 privremena traka 4x50m")
    mask_me3 = df_fakture["Materijal"].map(norm_text).eq(me508_3)
    mask_me4 = df_fakture["Materijal"].map(norm_text).eq(me508_4)
    df_fakture.loc[mask_me3 | mask_me4, "Jedinica mere za lager - skidanje količine"] = "kom"

    # 3) injekt map
    broj_pakera_map = _broj_pakera_po_setu(df_fakture)

    # 4) prvi prolaz
    df_pre = df_fakture.copy()
    df_pre[["Količina za skidanje sa lagera", "Napomena konverzije"]] = df_pre.apply(
        lambda row: calc_skidanje(row, broj_pakera_map, rules_dict),
        axis=1
    )
    neuspeh_pre = df_pre[df_pre["Količina za skidanje sa lagera"].isna()].copy()

    # 5) heuristike (bazirano na rules_dict)
    rules_all = {k: v[:] for k, v in rules_dict.items()}
    for _, r in neuspeh_pre[["Materijal", "Jedinica mere za fakturisanje", "Jedinica mere za lager - skidanje količine"]].drop_duplicates().iterrows():
        rr = make_heur_rule(r["Jedinica mere za fakturisanje"], r["Jedinica mere za lager - skidanje količine"])
        if rr is not None:
            rules_all.setdefault(canon_mat(r["Materijal"]), []).append(rr)

    # 6) drugi prolaz
    df_posle = df_fakture.copy()
    df_posle[["Količina za skidanje sa lagera", "Napomena konverzije"]] = df_posle.apply(
        lambda row: calc_skidanje(row, broj_pakera_map, rules_all),
        axis=1
    )

    # 7) PB2K spec m2/m -> lit
    PB2K_NAME_NORM = norm_text("HYPERDESMO PB 2K A+B, 20+20lit")
    materijal_norm = df_posle["Materijal"].map(norm_text)
    jm_fakt_norm = df_posle["Jedinica mere za fakturisanje"].map(norm_unit)
    jm_lager_norm = df_posle["Jedinica mere za lager - skidanje količine"].map(norm_unit)

    mask_pb2k = materijal_norm.eq(PB2K_NAME_NORM)
    mask_pb2k_m2 = mask_pb2k & jm_fakt_norm.eq("m2") & jm_lager_norm.eq("lit")
    df_posle.loc[mask_pb2k_m2, "Količina za skidanje sa lagera"] = (
        pd.to_numeric(df_posle.loc[mask_pb2k_m2, "Količina za fakturisanje"], errors="coerce") * PB2K_LIT_PER_M2
    )
    mask_pb2k_m = mask_pb2k & jm_fakt_norm.eq("m") & jm_lager_norm.eq("lit")
    df_posle.loc[mask_pb2k_m, "Količina za skidanje sa lagera"] = (
        pd.to_numeric(df_posle.loc[mask_pb2k_m, "Količina za fakturisanje"], errors="coerce") * PB2K_LIT_PER_M2
    )

    # kom -> round
    mask_kom_final = jm_lager_norm.eq("kom")
    df_posle.loc[mask_kom_final, "Količina za skidanje sa lagera"] = pd.to_numeric(
        df_posle.loc[mask_kom_final, "Količina za skidanje sa lagera"], errors="coerce"
    ).round(0)

    # 8) uporedba
    df_posle["Količina za skidanje sa lagera"] = pd.to_numeric(df_posle["Količina za skidanje sa lagera"], errors="coerce")

    potrosnja = (
        df_posle.groupby("Materijal", dropna=False)["Količina za skidanje sa lagera"]
        .sum(min_count=1)
        .reset_index()
        .rename(columns={"Količina za skidanje sa lagera": "Ukupna_potrošnja"})
    )

    lager_stanje = df_uvoz.iloc[1].to_frame().reset_index()
    lager_stanje.columns = ["Materijal", "Stanje_na_lageru"]
    lager_stanje["Stanje_na_lageru"] = pd.to_numeric(lager_stanje["Stanje_na_lageru"], errors="coerce")

    uporedba = pd.merge(potrosnja, lager_stanje, on="Materijal", how="outer")
    uporedba["Razlika_pre"] = uporedba["Stanje_na_lageru"] - uporedba["Ukupna_potrošnja"]

    # (tvoj postojeći dijagnostički koef)
    uporedba[["Koeficijent", "Coef_min", "Coef_max", "Flex", "Napomena_coef"]] = uporedba.apply(_calc_coef, axis=1)

    # ======================================================
    # 9) POPIS -> Koef_popis + Finalna_potrošnja
    # ======================================================
    if stvarno_file is not None:
        df_stvarno = parse_stvarno_stanje_excel(stvarno_file, value_row_idx=1)
        uporedba = uporedba.merge(df_stvarno, on="Materijal", how="left")

        # Koef_popis = knjigovodstvo / stvarno
        uporedba["Koef_popis"] = np.where(
            (uporedba["Stanje_na_magacinu"].notna()) &
            (uporedba["Stanje_na_magacinu"] > 0) &
            (uporedba["Stanje_na_lageru"].notna()),
            uporedba["Stanje_na_lageru"] / uporedba["Stanje_na_magacinu"],
            np.nan
        )

        uporedba["Finalna_potrošnja"] = uporedba["Ukupna_potrošnja"] * uporedba["Koef_popis"]
    else:
        uporedba["Stanje_na_magacinu"] = np.nan
        uporedba["Koef_popis"] = np.nan
        uporedba["Finalna_potrošnja"] = np.nan

    # ekstremi (po Napomena_coef, ali sa novim kolonama)
    ekstremni = uporedba[uporedba["Napomena_coef"].str.contains("EKSTREMNO", na=False)].copy()

    # pregled pravila koja su korišćena (za tab)
    rules_used_df = rules_to_df(rules_dict)

    return uporedba, ekstremni, df_posle, rules_used_df


# ======================================================
# 13) Pregled pravila – (read-only)
# ======================================================

def get_rules_overview():
    rows = []
    for mat, rules in material_rules.items():
        for idx, r in enumerate(rules, start=1):
            desc = getattr(r, "_rule_desc", r.__name__)
            rows.append({"Materijal (norm)": mat, "Rule #": idx, "Opis pravila": desc})
    return pd.DataFrame(rows)


# ======================================================
# 14) Export
# ======================================================

def export_to_excel(uporedba, ekstremni, df_fakture_posle):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        uporedba.to_excel(writer, index=False, sheet_name="koeficijenti_i_uporedba")
        ekstremni.to_excel(writer, index=False, sheet_name="ekstremni")
        df_fakture_posle.to_excel(writer, index=False, sheet_name="fakture_obracun")
    buffer.seek(0)
    return buffer

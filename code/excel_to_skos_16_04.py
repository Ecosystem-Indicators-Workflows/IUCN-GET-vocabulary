"""
Excel to SKOS/RDF Converter (v4 — external translation CSVs + placeholder defs)
================================================================================
Converts the IUCN GET profiles workbook into SKOS/RDF (Turtle + RDF/XML).

Changes vs. v3
--------------
* Accepts optional `--biome-translations <csv>` and `--realm-translations <csv>`
  for languages that aren't present in the Excel (es, fr). The CSV schema is:
      <code_col>, language, name, date
  where <code_col> is biome_code or realm_code (any column ending in "code"
  is accepted). Translations only ADD missing prefLabels — never overwrite an
  existing one from the Excel.
* New "Field descriptions" tab support: when an es/fr label is added from a
  translations CSV but no definition exists in that language, a placeholder
  definition is emitted (sourced from the Field descriptions tab) and tagged
  with skos:editorialNote = "Placeholder — translation pending"@<lang> so it
  can be located and replaced later.
* CLI moved to argparse.

Usage
-----
    pip install pandas openpyxl rdflib
    python excel_to_skos_v4.py <input.xlsx> \
        [--biome-translations biome-translations.csv] \
        [--realm-translations realm-translations.csv]
"""

import argparse
import os
import sys
import re
from collections import defaultdict

import pandas as pd
from rdflib import BNode, Graph, Literal, Namespace, URIRef
from rdflib.namespace import DCTERMS, OWL, RDF, RDFS, SKOS, XSD

# ── Namespaces ────────────────────────────────────────────────────────────────
BASE   = Namespace("https://w3id.org/iucn-get/")
IUCN   = Namespace("https://w3id.org/iucn-get/ontology#")
SCHEMA = Namespace("https://schema.org/")
FOAF   = Namespace("http://xmlns.com/foaf/0.1/")

TRANSLATABLE_LANGS = ("en", "es", "fr")
PLACEHOLDER_NOTE   = "Placeholder — translation pending"


# ── Literal helpers ───────────────────────────────────────────────────────────

def xsd_str(val):
    return Literal(str(val), datatype=XSD.string)

def xsd_uri(val):
    return Literal(str(val), datatype=XSD.anyURI)

def xsd_date(val):
    """Parse a date-ish value to xsd:date (YYYY-MM-DD)."""
    if val is None:
        return None
    try:
        d = pd.to_datetime(str(val), errors="coerce")
        if pd.isna(d):
            return None
        return Literal(d.strftime("%Y-%m-%d"), datatype=XSD.date)
    except Exception:
        return None

def lang_lit(val, lang):
    return Literal(str(val), lang=lang)


# ── General helpers ───────────────────────────────────────────────────────────

def slugify(text):
    text = str(text).strip()
    text = text.replace(".", "_")
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_-]+", "_", text)
    return text.strip("_")

def safe(val):
    if val is None or (hasattr(pd, "isna") and pd.isna(val)):
        return None
    s = str(val).strip()
    return s if s else None

def read_sheet(path, sheet):
    df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    return df

def col(df, *candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def strip_code_prefix(name, code):
    """Turn 'F1.1 Permanent upland streams' + code 'F1.1' → 'Permanent upland streams'."""
    if not name:
        return name
    pat = re.compile(r"^" + re.escape(code) + r"[\s.:\-]+")
    m = pat.match(name)
    return name[m.end():].strip() if m else name.strip()

def normalize_version(v):
    if not v:
        return None
    m = re.search(r"v\d+(?:\.\d+)*", v)
    return m.group(0) if m else v

def localized_url(url, lang):
    if not url or lang == "en":
        return url
    return re.sub(
        r"^(https?://[^/]+)/(?!(?:es|fr)/)",
        rf"\1/{lang}/",
        url,
    )

def realm_from_biome_code(bcode):
    """'F1' → 'F', 'FM1' → 'FM'."""
    if not bcode:
        return None
    m = re.match(r"^([A-Z]+)", bcode)
    return m.group(1) if m else None


# ── External-translation loaders ──────────────────────────────────────────────

def load_translations_csv(path):
    """Load a translations CSV into {(code, lang): {"name": ..., "date": ...}}.

    Expected columns (case-insensitive):
        - one column ending in "code" (e.g. biome_code, realm_code, code)
        - language (or lang)
        - name
        - date (optional)
    """
    if not path:
        return {}
    if not os.path.exists(path):
        print(f"  WARNING: translations file not found: {path}")
        return {}

    df = pd.read_csv(path, dtype=str)
    df.columns = [c.strip().lower() for c in df.columns]

    code_col = next(
        (c for c in df.columns if c == "code" or c.endswith("_code")),
        None,
    )
    lang_col = next((c for c in df.columns if c in ("language", "lang")), None)
    name_col = next((c for c in df.columns if c == "name"), None)
    date_col = next((c for c in df.columns if c in ("date", "date_modified", "modified")), None)

    if not (code_col and lang_col and name_col):
        print(f"  WARNING: {path} is missing required columns (code, language, name)")
        return {}

    out = {}
    for _, row in df.iterrows():
        code = safe(row[code_col])
        lang = safe(row[lang_col])
        name = safe(row[name_col])
        if not (code and lang and name):
            continue
        if lang not in TRANSLATABLE_LANGS:
            continue
        date = safe(row[date_col]) if date_col else None
        out[(code, lang)] = {"name": name, "date": date}
    print(f"  Loaded {len(out)} translation rows from {os.path.basename(path)}")
    return out


def load_field_descriptions(path):
    """Read the 'Field descriptions' tab into a dict {field_name_lower: description}.

    Layout assumed: a column identifying the field/concept ('field', 'name',
    'term', or 'concept') and a column with the description ('description',
    'definition', or 'field_description'). If the tab is missing or the
    columns can't be inferred, returns {}.
    """
    try:
        xl = pd.ExcelFile(path)
    except Exception as e:
        print(f"  WARNING: could not open workbook for field descriptions: {e}")
        return {}

    fd_sheet = next(
        (s for s in xl.sheet_names if "field" in s.lower() and "desc" in s.lower()),
        None,
    )
    if not fd_sheet:
        return {}

    df = read_sheet(path, fd_sheet)
    c_field = col(df, "field", "field_name", "name", "term", "concept")
    c_desc  = col(df, "description", "definition", "field_description")
    if not c_field or not c_desc:
        print(f"  WARNING: 'Field descriptions' tab present but could not infer columns "
              f"(saw {list(df.columns)})")
        return {}

    out = {}
    for _, row in df.iterrows():
        f = safe(row[c_field])
        d = safe(row[c_desc])
        if f and d:
            out[f.lower()] = d
    print(f"  Loaded {len(out)} field descriptions")
    return out


def placeholder_definition(field_descs, key_candidates):
    """Pick the first matching placeholder description for a class."""
    if not field_descs:
        return None
    for k in key_candidates:
        if k.lower() in field_descs:
            return field_descs[k.lower()]
    return None


# ── Concept registration ──────────────────────────────────────────────────────

def add_concept(g, uri, extra_type=None):
    g.add((uri, RDF.type, RDFS.Resource))
    g.add((uri, RDF.type, SKOS.Concept))
    if extra_type:
        g.add((uri, RDF.type, extra_type))
    g.add((uri, DCTERMS.identifier, xsd_uri(str(uri))))


def add_language_version(g, uri, lang, version, modified):
    bn = BNode()
    g.add((uri, IUCN["languageVersion"], bn))
    g.add((bn, DCTERMS.language, Literal(lang, datatype=XSD.language)))
    if version:
        g.add((bn, OWL.versionInfo, xsd_str(version)))
    if modified is not None:
        g.add((bn, DCTERMS.modified, modified))
    return bn


# ── Ontology header: class definitions ────────────────────────────────────────

CLASS_DEFS = {
    "Realm": {
        "label": "Realm",
        "comment": (
            "One of five major components of the biosphere that differ "
            "fundamentally in ecosystem organisation and function: terrestrial, "
            "freshwater, marine, subterranean, atmospheric and combinations of "
            "these (transitional realms). Because variation in nature is "
            "continuous, we also include transitional realms, where the realms "
            "meet and have their own unique organisation and function."
        ),
        "source": "https://global-ecosystems.org/",
    },
    "Biome": {
        "label": "Functional biome",
        "comment": (
            "A component of a realm united by broad features of ecosystem "
            "structure and one or a few common major ecological drivers that "
            "regulate major ecological functions, derived from the top-down by "
            "subdivision of realms (level 1)."
        ),
        "scope_note": (
            "In GET, biome is different from the use of biome in other "
            "contexts — it is defined functionally rather than by vegetation "
            "physiognomy or climate alone."
        ),
        "source": "https://global-ecosystems.org/",
    },
    "EcosystemFunctionalGroup": {
        "label": "Ecosystem Functional Group",
        "comment": (
            "A group of related ecosystems within a biome that share common "
            "ecological drivers, which in turn promote similar biotic traits "
            "that characterise the group. Derived from the top-down by "
            "subdivision of biomes."
        ),
        "alt_label": "EFG",
        "source": "https://global-ecosystems.org/",
    },
}

def add_class_definitions(g):
    for local, d in CLASS_DEFS.items():
        cls = IUCN[local]
        g.add((cls, RDF.type, RDFS.Class))
        g.add((cls, RDFS.label, lang_lit(d["label"], "en")))
        g.add((cls, RDFS.comment, lang_lit(d["comment"], "en")))
        if d.get("alt_label"):
            g.add((cls, SKOS.altLabel, lang_lit(d["alt_label"], "en")))
        if d.get("scope_note"):
            g.add((cls, SKOS.scopeNote, lang_lit(d["scope_note"], "en")))
        if d.get("source"):
            g.add((cls, DCTERMS.source, URIRef(d["source"])))


# ── Ontology header: property definitions ────────────────────────────────────

PROPERTY_DEFS = {
    "realmCode": {
        "label": "realm code",
        "comment": "Code assigned to a Realm in the IUCN Global Ecosystem Typology.",
        "domain": IUCN["Realm"], "range": XSD.string,
    },
    "realmType": {
        "label": "realm type",
        "comment": "Type designation of the Realm (e.g. core or transitional).",
        "domain": IUCN["Realm"], "range": XSD.string,
    },
    "biomeCode": {
        "label": "biome code",
        "comment": "Code assigned to a Biome in the IUCN Global Ecosystem Typology.",
        "domain": IUCN["Biome"], "range": XSD.string,
    },
    "efgCode": {
        "label": "ecosystem functional group code",
        "comment": "Code assigned to an Ecosystem Functional Group in the IUCN Global Ecosystem Typology.",
        "domain": IUCN["EcosystemFunctionalGroup"], "range": XSD.string,
    },
    "keyFeatures": {
        "label": "key features",
        "comment": "Characteristic ecological and functional features of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"], "range": RDFS.Literal,
    },
    "distributionSummary": {
        "label": "distribution summary",
        "comment": "Summary description of the distribution of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"], "range": RDFS.Literal,
    },
    "ecosystemProperties": {
        "label": "ecosystem properties",
        "comment": "Description of notable structural or functional properties of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"], "range": RDFS.Literal,
    },
    "ecologicalDrivers": {
        "label": "ecological drivers",
        "comment": "Description of the main ecological drivers shaping the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"], "range": RDFS.Literal,
    },
    "distribution": {
        "label": "distribution",
        "comment": "Detailed description of the geographic or environmental distribution of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"], "range": RDFS.Literal,
        "subproperty_of": SKOS.note,
    },
    "languageVersion": {
        "label": "language version",
        "comment": (
            "Per-language version metadata (language, version tag, modification "
            "date). Used to track that translations of a concept may have been "
            "modified independently."
        ),
        "range": RDFS.Resource,
    },
    "latestUpdate": {
        "label": "latest update",
        "comment": (
            "Points to the iucn:languageVersion blank node carrying the most "
            "recent dcterms:modified date for this concept."
        ),
        "range": RDFS.Resource,
    },
    "latestModificationLanguage": {
        "label": "latest modification language",
        "comment": (
            "Language code (ISO 639-1) of the translation whose modification "
            "date was used as the top-level dcterms:modified value."
        ),
        "range": XSD.language,
    },
}

def add_property_definitions(g):
    for local, d in PROPERTY_DEFS.items():
        p = IUCN[local]
        g.add((p, RDF.type, RDF.Property))
        g.add((p, RDFS.label, lang_lit(d["label"], "en")))
        g.add((p, RDFS.comment, lang_lit(d["comment"], "en")))
        if d.get("domain"):
            g.add((p, RDFS.domain, d["domain"]))
        if d.get("range"):
            g.add((p, RDFS.range, d["range"]))
        if d.get("subproperty_of"):
            g.add((p, RDFS.subPropertyOf, d["subproperty_of"]))


# ── Per-concept emission helpers ──────────────────────────────────────────────

def emit_translatable(g, uri, lang, fields):
    for pred, val in fields:
        if val:
            g.add((uri, pred, lang_lit(val, lang)))

def get_pref_label(g, uri, lang):
    for o in g.objects(uri, SKOS.prefLabel):
        if isinstance(o, Literal) and o.language == lang:
            return str(o)
    return None

def has_definition_in_lang(g, uri, lang):
    for o in g.objects(uri, SKOS.definition):
        if isinstance(o, Literal) and o.language == lang:
            return True
    return False


def merge_translation(g, uri, code, lang, name, placeholder_def):
    """Add prefLabel + (optional placeholder) definition for one (concept, lang).
    Never overwrites existing labels/definitions in the same language."""
    if not get_pref_label(g, uri, lang):
        g.add((uri, SKOS.prefLabel, lang_lit(strip_code_prefix(name, code), lang)))
    if placeholder_def and not has_definition_in_lang(g, uri, lang):
        g.add((uri, SKOS.definition, lang_lit(placeholder_def, lang)))
        g.add((uri, SKOS.editorialNote, lang_lit(PLACEHOLDER_NOTE, lang)))


# ── Build graph ───────────────────────────────────────────────────────────────

def build_graph(path, biome_trans=None, realm_trans=None, field_descs=None):
    biome_trans = biome_trans or {}
    realm_trans = realm_trans or {}
    field_descs = field_descs or {}

    g = Graph()
    g.bind("rdf",     RDF)
    g.bind("rdfs",    RDFS)
    g.bind("xsd",     XSD)
    g.bind("skos",    SKOS)
    g.bind("dcterms", DCTERMS)
    g.bind("owl",     OWL)
    g.bind("iucn",    IUCN)
    g.bind("base",    BASE)
    g.bind("schema",  SCHEMA)
    g.bind("foaf",    FOAF)

    add_class_definitions(g)
    add_property_definitions(g)

    # Concept scheme
    scheme = BASE["IUCNEcosystemTypology"]
    g.add((scheme, RDF.type, SKOS.ConceptScheme))
    g.add((scheme, DCTERMS.identifier, xsd_uri(str(scheme))))
    g.add((scheme, SKOS.prefLabel, lang_lit("IUCN Global Ecosystem Typology", "en")))
    g.add((scheme, DCTERMS.description,
           lang_lit("A hierarchical classification of Earth's ecosystems.", "en")))

    xl = pd.ExcelFile(path)
    sheet_names = xl.sheet_names
    print(f"Found sheets: {sheet_names}")

    realm_placeholder = placeholder_definition(field_descs, ["realm", "realm description", "realm definition"])
    biome_placeholder = placeholder_definition(field_descs, ["biome", "functional biome", "biome description", "biome definition"])

    # ── Level 1: Realms ──────────────────────────────────────────────────────
    realm_sheet = next((s for s in sheet_names if "realm" in s.lower()), None)
    if realm_sheet:
        df = read_sheet(path, realm_sheet)
        print(f"  Level 1 Realms — columns: {list(df.columns)}")
        c_code = col(df, "realm_code", "coderealm", "code")
        c_name = col(df, "name", "realm")
        c_type = col(df, "realm_type", "type")
        c_desc = col(df, "description", "typedescription")

        for _, row in df.iterrows():
            code = safe(row[c_code]) if c_code else None
            if not code:
                continue
            name  = safe(row[c_name]) if c_name else None
            rtype = safe(row[c_type]) if c_type else None
            desc  = safe(row[c_desc]) if c_desc else None

            uri = BASE[f"Realm_{slugify(code)}"]
            add_concept(g, uri, IUCN["Realm"])
            g.add((uri, SKOS.inScheme,        scheme))
            g.add((uri, SKOS.topConceptOf,    scheme))
            g.add((scheme, SKOS.hasTopConcept, uri))
            g.add((uri, SKOS.notation,        xsd_str(code)))
            g.add((uri, IUCN["realmCode"],    xsd_str(code)))
            if rtype:
                g.add((uri, IUCN["realmType"], xsd_str(rtype)))
            if name:
                g.add((uri, SKOS.prefLabel, lang_lit(strip_code_prefix(name, code), "en")))
            if desc:
                g.add((uri, SKOS.definition, lang_lit(desc, "en")))

        # Merge realm translations from external CSV (es, fr, …)
        if realm_trans:
            print(f"  Merging {len(realm_trans)} realm translations from CSV")
            for (rcode, lang), payload in realm_trans.items():
                uri = BASE[f"Realm_{slugify(rcode)}"]
                if (uri, RDF.type, SKOS.Concept) not in g:
                    print(f"    skipping unknown realm code: {rcode}")
                    continue
                merge_translation(g, uri, rcode, lang, payload["name"], realm_placeholder)
    else:
        print("  WARNING: No Realms sheet found.")

    # ── Level 2: Biomes ──────────────────────────────────────────────────────
    biome_sheet = next((s for s in sheet_names if "biome" in s.lower()), None)
    if biome_sheet:
        df = read_sheet(path, biome_sheet)
        print(f"  Level 2 Biomes — columns: {list(df.columns)}")
        c_bcode   = col(df, "biome_code", "biome", "code")
        c_rcode   = col(df, "realm_code", "codenamerealm", "realm")
        c_name    = col(df, "name", "biome_name")
        c_desc    = col(df, "description", "codedescription")
        c_contrib = col(df, "contributors", "contributor")
        c_lang    = col(df, "lang", "language")
        c_ver     = col(df, "version")
        c_dmod    = col(df, "date_modified", "modified")

        grouped = defaultdict(list)
        for _, row in df.iterrows():
            bcode = safe(row[c_bcode]) if c_bcode else None
            if bcode:
                grouped[bcode].append(row)

        for bcode, rows in grouped.items():
            uri = BASE[f"Biome_{slugify(bcode)}"]
            add_concept(g, uri, IUCN["Biome"])
            g.add((uri, SKOS.inScheme,    scheme))
            g.add((uri, SKOS.notation,    xsd_str(bcode)))
            g.add((uri, IUCN["biomeCode"], xsd_str(bcode)))

            rcode = safe(rows[0][c_rcode]) if c_rcode else None
            if rcode:
                realm_uri = BASE[f"Realm_{slugify(rcode)}"]
                g.add((uri, SKOS.broader, realm_uri))
                g.add((realm_uri, SKOS.narrower, uri))

            contrib = safe(rows[0][c_contrib]) if c_contrib else None
            if contrib:
                g.add((uri, DCTERMS.contributor, xsd_str(contrib)))

            latest_modified = None
            latest_bnode = None
            latest_lang = None
            ordered = sorted(rows, key=lambda r: TRANSLATABLE_LANGS.index(
                safe(r[c_lang]) if c_lang and safe(r[c_lang]) in TRANSLATABLE_LANGS
                else "en"
            ))

            for row in ordered:
                lang = safe(row[c_lang]) if c_lang else "en"
                if lang not in TRANSLATABLE_LANGS:
                    lang = "en"

                name = safe(row[c_name]) if c_name else None
                desc = safe(row[c_desc]) if c_desc else None
                if name:
                    label = strip_code_prefix(name, bcode)
                    # Skip non-English prefLabel if it duplicates the English one —
                    # that's an untranslated placeholder in the workbook, not a real
                    # translation. Leaving the slot empty lets the CSV merge fill it.
                    en_pref = get_pref_label(g, uri, "en")
                    if lang == "en" or label != en_pref:
                        g.add((uri, SKOS.prefLabel, lang_lit(label, lang)))
                if desc:
                    g.add((uri, SKOS.definition, lang_lit(desc, lang)))

                ver  = normalize_version(safe(row[c_ver])) if c_ver else None
                dmod = xsd_date(safe(row[c_dmod])) if c_dmod else None
                bn = add_language_version(g, uri, lang, ver, dmod)

                if dmod is not None:
                    if latest_modified is None or str(dmod) > str(latest_modified):
                        latest_modified = dmod
                        latest_bnode = bn
                        latest_lang = lang

            if latest_modified is not None:
                g.add((uri, DCTERMS.modified, latest_modified))
            if latest_bnode is not None:
                g.add((uri, IUCN["latestUpdate"], latest_bnode))
            if latest_lang is not None:
                g.add((uri, IUCN["latestModificationLanguage"],
                       Literal(latest_lang, datatype=XSD.language)))

        # Merge biome translations from external CSV
        if biome_trans:
            print(f"  Merging {len(biome_trans)} biome translations from CSV")
            for (bcode, lang), payload in biome_trans.items():
                uri = BASE[f"Biome_{slugify(bcode)}"]
                if (uri, RDF.type, SKOS.Concept) not in g:
                    print(f"    skipping unknown biome code: {bcode}")
                    continue
                merge_translation(g, uri, bcode, lang, payload["name"], biome_placeholder)

                # Track the translation date as a per-language version blank node
                dmod = xsd_date(payload.get("date"))
                if dmod is not None:
                    add_language_version(g, uri, lang, None, dmod)
    else:
        print("  WARNING: No Biomes sheet found.")

    # ── Level 3: EFGs ────────────────────────────────────────────────────────
    efg_by_biome = defaultdict(list)

    efg_sheet = next(
        (s for s in sheet_names
         if "ecosystem" in s.lower() and "functional" in s.lower()),
        None,
    )
    if efg_sheet:
        df = read_sheet(path, efg_sheet)
        print(f"  Level 3 EFG — columns: {list(df.columns)}")
        c_code    = col(df, "code")
        c_name    = col(df, "name")
        c_props   = col(df, "ecosystem_properties", "properties")
        c_drivers = col(df, "ecological_drivers", "drivers")
        c_dist    = col(df, "distribution")
        c_contrib = col(df, "contributors", "contributor")
        c_lang    = col(df, "lang", "language")
        c_ver     = col(df, "version")
        c_dmod    = col(df, "date_modified", "modified")

        grouped = defaultdict(list)
        for _, row in df.iterrows():
            code = safe(row[c_code]) if c_code else None
            if code:
                grouped[code].append(row)

        for code, rows in grouped.items():
            uri = BASE[f"EFG_{slugify(code)}"]
            add_concept(g, uri, IUCN["EcosystemFunctionalGroup"])
            g.add((uri, SKOS.inScheme, scheme))
            g.add((uri, SKOS.notation, xsd_str(code)))
            g.add((uri, IUCN["efgCode"], xsd_str(code)))

            parts = code.split(".")
            if len(parts) >= 2:
                bcode = parts[0]
                biome_uri = BASE[f"Biome_{slugify(bcode)}"]
                g.add((uri, SKOS.broader, biome_uri))
                g.add((biome_uri, SKOS.narrower, uri))
                efg_by_biome[bcode].append(uri)

                rcode = realm_from_biome_code(bcode)
                if rcode:
                    realm_uri = BASE[f"Realm_{slugify(rcode)}"]
                    g.add((uri, SKOS.broaderTransitive, realm_uri))
                    g.add((realm_uri, SKOS.narrowerTransitive, uri))

            contrib = safe(rows[0][c_contrib]) if c_contrib else None
            if contrib:
                g.add((uri, DCTERMS.contributor, xsd_str(contrib)))

            latest_modified = None
            latest_bnode = None
            latest_lang = None
            ordered = sorted(rows, key=lambda r: TRANSLATABLE_LANGS.index(
                safe(r[c_lang]) if c_lang and safe(r[c_lang]) in TRANSLATABLE_LANGS
                else "en"
            ))

            for row in ordered:
                lang = safe(row[c_lang]) if c_lang else "en"
                if lang not in TRANSLATABLE_LANGS:
                    lang = "en"

                name    = safe(row[c_name])    if c_name    else None
                props   = safe(row[c_props])   if c_props   else None
                drivers = safe(row[c_drivers]) if c_drivers else None
                dist    = safe(row[c_dist])    if c_dist    else None

                if name:
                    label = strip_code_prefix(name, code)
                    en_pref = get_pref_label(g, uri, "en")
                    if lang == "en" or label != en_pref:
                        g.add((uri, SKOS.prefLabel, lang_lit(label, lang)))
                emit_translatable(g, uri, lang, [
                    (IUCN["ecosystemProperties"], props),
                    (IUCN["ecologicalDrivers"],   drivers),
                    (IUCN["distribution"],        dist),
                ])

                ver  = normalize_version(safe(row[c_ver])) if c_ver else None
                dmod = xsd_date(safe(row[c_dmod])) if c_dmod else None
                bn = add_language_version(g, uri, lang, ver, dmod)

                if dmod is not None:
                    if latest_modified is None or str(dmod) > str(latest_modified):
                        latest_modified = dmod
                        latest_bnode = bn
                        latest_lang = lang

            if latest_modified is not None:
                g.add((uri, DCTERMS.modified, latest_modified))
            if latest_bnode is not None:
                g.add((uri, IUCN["latestUpdate"], latest_bnode))
            if latest_lang is not None:
                g.add((uri, IUCN["latestModificationLanguage"],
                       Literal(latest_lang, datatype=XSD.language)))

        for bcode, uris in efg_by_biome.items():
            for i, a in enumerate(uris):
                for b in uris[i + 1:]:
                    g.add((a, SKOS.related, b))
                    g.add((b, SKOS.related, a))
    else:
        print("  WARNING: No EFG sheet found.")

    # ── Level 3: Short Descriptions ──────────────────────────────────────────
    short_sheet = next(
        (s for s in sheet_names if "short" in s.lower() and "desc" in s.lower()),
        None,
    )
    if short_sheet:
        df = read_sheet(path, short_sheet)
        print(f"  Level 3 Short Desc — columns: {list(df.columns)}")
        c_code    = col(df, "code")
        c_bcode   = col(df, "biome_code", "biome")
        c_name    = col(df, "name")
        c_sname   = col(df, "short_name")
        c_dcre    = col(df, "date_created", "created")
        c_sdesc   = col(df, "short_description", "short_desc")
        c_feat    = col(df, "key_features")
        c_distsum = col(df, "distribution_summary")
        c_url     = col(df, "url")

        for _, row in df.iterrows():
            code = safe(row[c_code]) if c_code else None
            if not code:
                continue

            uri = BASE[f"EFG_{slugify(code)}"]
            if (uri, RDF.type, SKOS.Concept) not in g:
                add_concept(g, uri, IUCN["EcosystemFunctionalGroup"])
                g.add((uri, SKOS.inScheme, scheme))
                g.add((uri, SKOS.notation, xsd_str(code)))
                g.add((uri, IUCN["efgCode"], xsd_str(code)))

            name    = safe(row[c_name])    if c_name    else None
            sname   = safe(row[c_sname])   if c_sname   else None
            sdesc   = safe(row[c_sdesc])   if c_sdesc   else None
            feat    = safe(row[c_feat])    if c_feat    else None
            distsum = safe(row[c_distsum]) if c_distsum else None
            dcre    = xsd_date(safe(row[c_dcre])) if c_dcre else None
            url     = safe(row[c_url])     if c_url     else None
            bcode   = safe(row[c_bcode])   if c_bcode   else None

            if name and not list(g.triples((uri, SKOS.prefLabel, None))):
                g.add((uri, SKOS.prefLabel,
                       lang_lit(strip_code_prefix(name, code), "en")))

            if sname:
                short_label = strip_code_prefix(sname, code)
                en_pref = get_pref_label(g, uri, "en")
                if short_label and short_label != en_pref:
                    g.add((uri, SKOS.altLabel, lang_lit(short_label, "en")))

            if sdesc:
                g.add((uri, SKOS.scopeNote, lang_lit(sdesc, "en")))
            if feat:
                g.add((uri, IUCN["keyFeatures"], lang_lit(feat, "en")))
            if distsum:
                g.add((uri, IUCN["distributionSummary"], lang_lit(distsum, "en")))
            if dcre is not None:
                g.add((uri, DCTERMS.created, dcre))

            if url:
                for lang in TRANSLATABLE_LANGS:
                    g.add((uri, FOAF["page"], URIRef(localized_url(url, lang))))

            if bcode and not list(g.triples((uri, SKOS.broader, None))):
                biome_uri = BASE[f"Biome_{slugify(bcode)}"]
                g.add((uri, SKOS.broader, biome_uri))
                g.add((biome_uri, SKOS.narrower, uri))
    else:
        print("  WARNING: No Short Descriptions sheet found.")

    return g


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Convert IUCN GET Excel workbook to SKOS/RDF."
    )
    parser.add_argument("xlsx", help="Path to the IUCN GET workbook (.xlsx)")
    parser.add_argument(
        "--biome-translations",
        default=None,
        help="Optional CSV with biome name translations (columns: biome_code, language, name, date)",
    )
    parser.add_argument(
        "--realm-translations",
        default=None,
        help="Optional CSV with realm name translations (columns: realm_code, language, name, date)",
    )
    args = parser.parse_args()

    path = args.xlsx
    in_dir, in_name = os.path.split(os.path.abspath(path))
    stem = in_name.rsplit(".", 1)[0]

    out_dir = os.path.join(os.path.dirname(in_dir), "output")
    os.makedirs(out_dir, exist_ok=True)

    print(f"\nReading: {path}")
    biome_trans = load_translations_csv(args.biome_translations)
    realm_trans = load_translations_csv(args.realm_translations)
    field_descs = load_field_descriptions(path)

    g = build_graph(
        path,
        biome_trans=biome_trans,
        realm_trans=realm_trans,
        field_descs=field_descs,
    )

    ttl_path = os.path.join(out_dir, f"{stem}_skos.ttl")
    rdf_path = os.path.join(out_dir, f"{stem}_skos.rdf")
    g.serialize(destination=ttl_path, format="turtle")
    g.serialize(destination=rdf_path, format="xml")

    print(f"\nDone.")
    print(f"  Turtle  -> {ttl_path}")
    print(f"  RDF/XML -> {rdf_path}")
    print(f"  Triples : {len(g)}")


if __name__ == "__main__":
    main()
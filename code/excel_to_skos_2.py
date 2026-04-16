"""
Excel to SKOS/RDF Converter (v2 — multilingual)
================================================
Converts the IUCN GET profiles workbook into SKOS/RDF (Turtle + RDF/XML).

Changes vs. v1
--------------
* Handles multilingual sheets with a `lang` column (en/es/fr). Translatable
  fields (prefLabel, definition, ecosystemProperties, ecologicalDrivers,
  distribution) are emitted as language-tagged literals.
* Per-language modification dates and version info are kept on a blank-node
  `iucn:languageVersion` so each translation carries its own provenance.
  The top-level `dcterms:modified` is set to the most recent date across
  all languages.
* Splits the human-readable name into `skos:prefLabel` (lang-tagged) and
  `skos:notation` (the code, e.g. "F1.1").
* Dates are emitted as `xsd:date` (date only).
* `owl:versionInfo` is a short tag like "v2.1".
* Class declarations for iucn:Realm / iucn:Biome / iucn:EcosystemFunctionalGroup
  are added with the official definitions.
* Per-language URLs are built from the English URL by inserting /es/ or /fr/.

Usage
-----
    pip install pandas openpyxl rdflib
    python excel_to_skos_rdf.py <input.xlsx>
"""

import sys
import re
from collections import defaultdict

import pandas as pd
from rdflib import BNode, Graph, Literal, Namespace, URIRef
from rdflib.namespace import DCTERMS, OWL, RDF, RDFS, SKOS, XSD

# ── Namespaces ────────────────────────────────────────────────────────────────
BASE   = Namespace("https://example.org/iucn/")
IUCN   = Namespace("https://example.org/iucn/ontology#")
SCHEMA = Namespace("https://schema.org/")

TRANSLATABLE_LANGS = ("en", "es", "fr")


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
    # Replace dots first so "F1.1" → "F1_1" (not "F11")
    text = text.replace(".", "_")
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_-]+", "_", text)
    return text.strip("_")

def safe(val):
    if val is None or pd.isna(val):
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
    """Turn 'F1.1 Permanent upland streams' + code 'F1.1' → 'Permanent upland streams'.
    Also handles 'F1. Rivers and streams biome' + 'F1' → 'Rivers and streams biome'."""
    if not name:
        return name
    pat = re.compile(r"^" + re.escape(code) + r"\.?\s*[:.\-]?\s*")
    return pat.sub("", name).strip()

def normalize_version(v):
    if not v:
        return None
    m = re.search(r"v\d+(?:\.\d+)*", v)
    return m.group(0) if m else v

def localized_url(url, lang):
    """Insert /es/ or /fr/ into a global-ecosystems.org URL for non-en languages."""
    if not url or lang == "en":
        return url
    return re.sub(
        r"^(https?://[^/]+)/(?!(?:es|fr)/)",
        rf"\1/{lang}/",
        url,
    )


# ── Concept registration ──────────────────────────────────────────────────────

def add_concept(g, uri, extra_type=None):
    g.add((uri, RDF.type, RDFS.Resource))
    g.add((uri, RDF.type, SKOS.Concept))
    if extra_type:
        g.add((uri, RDF.type, extra_type))
    # Canonical identifier plus one per language
    g.add((uri, DCTERMS.identifier, xsd_uri(str(uri))))
    local = str(uri).rsplit("/", 1)[-1]
    for lang in TRANSLATABLE_LANGS:
        g.add((uri, DCTERMS.identifier,
               xsd_uri(f"{BASE}{lang}/{local}")))


def add_language_version(g, uri, lang, version, modified, url=None):
    """Attach per-language provenance as a blank node."""
    bn = BNode()
    g.add((uri, IUCN["languageVersion"], bn))
    g.add((bn, DCTERMS.language, Literal(lang)))
    if version:
        g.add((bn, OWL.versionInfo, xsd_str(version)))
    if modified is not None:
        g.add((bn, DCTERMS.modified, modified))
    if url:
        g.add((bn, SCHEMA["url"], URIRef(url)))


# ── Ontology header (class definitions) ───────────────────────────────────────

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


# ── Per-concept emission helper for multilingual rows ─────────────────────────

def emit_translatable(g, uri, lang, fields):
    """fields: list of (predicate, value)."""
    for pred, val in fields:
        if val:
            g.add((uri, pred, lang_lit(val, lang)))


# ── Build graph ───────────────────────────────────────────────────────────────

def build_graph(path):
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

    add_class_definitions(g)

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

    # ── Level 1: Realms (single language, English) ───────────────────────────
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
            g.add((uri, SKOS.inScheme,     scheme))
            g.add((uri, SKOS.topConceptOf, scheme))
            g.add((uri, SKOS.notation,     xsd_str(code)))
            g.add((uri, IUCN["realmCode"], xsd_str(code)))
            if rtype:
                g.add((uri, IUCN["realmType"], xsd_str(rtype)))
            if name:
                g.add((uri, SKOS.prefLabel, lang_lit(strip_code_prefix(name, code), "en")))
            if desc:
                g.add((uri, SKOS.definition, lang_lit(desc, "en")))
    else:
        print("  WARNING: No Realms sheet found.")

    # ── Level 2: Biomes (multilingual) ────────────────────────────────────────
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

        # Group rows by biome code
        grouped = defaultdict(list)
        for _, row in df.iterrows():
            bcode = safe(row[c_bcode]) if c_bcode else None
            if bcode:
                grouped[bcode].append(row)

        for bcode, rows in grouped.items():
            uri = BASE[f"Biome_{slugify(bcode)}"]
            add_concept(g, uri, IUCN["Biome"])
            g.add((uri, SKOS.inScheme, scheme))
            g.add((uri, SKOS.notation, xsd_str(bcode)))
            g.add((uri, IUCN["biomeCode"], xsd_str(bcode)))

            # Realm link (same across all language rows)
            rcode = safe(rows[0][c_rcode]) if c_rcode else None
            if rcode:
                realm_uri = BASE[f"Realm_{slugify(rcode)}"]
                g.add((uri, SKOS.broader, realm_uri))
                g.add((realm_uri, SKOS.narrower, uri))

            # Contributor (same across rows)
            contrib = safe(rows[0][c_contrib]) if c_contrib else None
            if contrib:
                g.add((uri, DCTERMS.contributor, xsd_str(contrib)))

            latest_modified = None
            for row in rows:
                lang = safe(row[c_lang]) if c_lang else "en"
                if lang not in TRANSLATABLE_LANGS:
                    lang = "en"

                name = safe(row[c_name]) if c_name else None
                desc = safe(row[c_desc]) if c_desc else None
                if name:
                    g.add((uri, SKOS.prefLabel,
                           lang_lit(strip_code_prefix(name, bcode), lang)))
                if desc:
                    g.add((uri, SKOS.definition, lang_lit(desc, lang)))

                ver  = normalize_version(safe(row[c_ver])) if c_ver else None
                dmod = xsd_date(safe(row[c_dmod])) if c_dmod else None
                add_language_version(g, uri, lang, ver, dmod)

                if dmod is not None:
                    if latest_modified is None or str(dmod) > str(latest_modified):
                        latest_modified = dmod

            if latest_modified is not None:
                g.add((uri, DCTERMS.modified, latest_modified))
    else:
        print("  WARNING: No Biomes sheet found.")

    # ── Level 3: EFGs (multilingual) ─────────────────────────────────────────
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

            # Biome link from code prefix (e.g. "F1.1" → "F1")
            parts = code.split(".")
            if len(parts) >= 2:
                biome_uri = BASE[f"Biome_{slugify(parts[0])}"]
                g.add((uri, SKOS.broader, biome_uri))
                g.add((biome_uri, SKOS.narrower, uri))

            contrib = safe(rows[0][c_contrib]) if c_contrib else None
            if contrib:
                g.add((uri, DCTERMS.contributor, xsd_str(contrib)))

            latest_modified = None
            for row in rows:
                lang = safe(row[c_lang]) if c_lang else "en"
                if lang not in TRANSLATABLE_LANGS:
                    lang = "en"

                name    = safe(row[c_name])    if c_name    else None
                props   = safe(row[c_props])   if c_props   else None
                drivers = safe(row[c_drivers]) if c_drivers else None
                dist    = safe(row[c_dist])    if c_dist    else None

                if name:
                    g.add((uri, SKOS.prefLabel,
                           lang_lit(strip_code_prefix(name, code), lang)))
                emit_translatable(g, uri, lang, [
                    (IUCN["ecosystemProperties"], props),
                    (IUCN["ecologicalDrivers"],   drivers),
                    (IUCN["distribution"],        dist),
                ])

                ver  = normalize_version(safe(row[c_ver])) if c_ver else None
                dmod = xsd_date(safe(row[c_dmod])) if c_dmod else None
                add_language_version(g, uri, lang, ver, dmod)

                if dmod is not None:
                    if latest_modified is None or str(dmod) > str(latest_modified):
                        latest_modified = dmod

            if latest_modified is not None:
                g.add((uri, DCTERMS.modified, latest_modified))
    else:
        print("  WARNING: No EFG sheet found.")

    # ── Level 3: Short Descriptions (English-only sheet, with URL) ───────────
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

            # prefLabel only if not already set from EFG sheet
            if name and not list(g.triples((uri, SKOS.prefLabel, None))):
                g.add((uri, SKOS.prefLabel,
                       lang_lit(strip_code_prefix(name, code), "en")))
            if sname:
                g.add((uri, SKOS.altLabel, lang_lit(sname, "en")))
            if sdesc:
                g.add((uri, SKOS.scopeNote, lang_lit(sdesc, "en")))
            if feat:
                g.add((uri, IUCN["keyFeatures"], lang_lit(feat, "en")))
            if distsum:
                g.add((uri, IUCN["distributionSummary"], lang_lit(distsum, "en")))
            if dcre is not None:
                g.add((uri, DCTERMS.created, dcre))

            # Per-language URLs
            if url:
                for lang in TRANSLATABLE_LANGS:
                    g.add((uri, SCHEMA["url"], URIRef(localized_url(url, lang))))

            if bcode and not list(g.triples((uri, SKOS.broader, None))):
                biome_uri = BASE[f"Biome_{slugify(bcode)}"]
                g.add((uri, SKOS.broader, biome_uri))
                g.add((biome_uri, SKOS.narrower, uri))
    else:
        print("  WARNING: No Short Descriptions sheet found.")

    return g


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print("Usage: python excel_to_skos_rdf.py <input.xlsx>")
        sys.exit(1)

    path = sys.argv[1]
    stem = path.rsplit(".", 1)[0]

    print(f"\nReading: {path}")
    g = build_graph(path)

    ttl_path = f"{stem}_skos.ttl"
    rdf_path = f"{stem}_skos.rdf"
    g.serialize(destination=ttl_path, format="turtle")
    g.serialize(destination=rdf_path, format="xml")

    print(f"\nDone.")
    print(f"  Turtle  -> {ttl_path}")
    print(f"  RDF/XML -> {rdf_path}")
    print(f"  Triples : {len(g)}")


if __name__ == "__main__":
    main()
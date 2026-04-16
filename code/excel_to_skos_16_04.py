"""
Excel to SKOS/RDF Converter (v3 — multilingual + w3id)
======================================================
Converts the IUCN GET profiles workbook into SKOS/RDF (Turtle + RDF/XML).

Changes vs. v2
--------------
* Base URI moved to https://w3id.org/iucn-get/ (canonical, persistent).
* Per-language identifiers removed — each concept has ONE canonical
  dcterms:identifier (https://w3id.org/iucn-get/EFG_F1_1). Language-specific
  web pages on global-ecosystems.org are linked via foaf:page instead,
  cleanly separating concept identity from language-specific representations.
* skos:altLabel is suppressed when equal to the prefLabel of the same lang.
* EFGs link to their realm via skos:broaderTransitive (biome stays on
  skos:broader as the direct parent).
* EFGs in the same biome are linked with skos:related.
* iucn:latestUpdate links a concept to its most recently modified
  language-version blank node.
* Custom iucn:* properties are now declared with rdfs:domain / rdfs:range
  in the ontology header (following William's example). iucn:distribution
  is declared as rdfs:subPropertyOf skos:note.

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
BASE   = Namespace("https://w3id.org/iucn-get/")
IUCN   = Namespace("https://w3id.org/iucn-get/ontology#")
SCHEMA = Namespace("https://schema.org/")
FOAF   = Namespace("http://xmlns.com/foaf/0.1/")

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
    Also handles 'F1. Rivers and streams biome' + 'F1' → 'Rivers and streams biome'.
    Critically, REQUIRES a separator (whitespace/./-/:) after the code, so that
    a name like 'Atmospheric' with code 'A' is left untouched (not chopped to
    'tmospheric')."""
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
    """Insert /es/ or /fr/ into a global-ecosystems.org URL for non-en languages."""
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


# ── Concept registration ──────────────────────────────────────────────────────

def add_concept(g, uri, extra_type=None):
    g.add((uri, RDF.type, RDFS.Resource))
    g.add((uri, RDF.type, SKOS.Concept))
    if extra_type:
        g.add((uri, RDF.type, extra_type))
    # Single canonical identifier — the concept is language-independent.
    # Language-specific web pages go on foaf:page instead.
    g.add((uri, DCTERMS.identifier, xsd_uri(str(uri))))


def add_language_version(g, uri, lang, version, modified):
    """Attach per-language provenance as a blank node. Returns the blank node."""
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


# ── Ontology header: property definitions (per William's email) ──────────────

PROPERTY_DEFS = {
    # ── Code properties ──
    "realmCode": {
        "label": "realm code",
        "comment": "Code assigned to a Realm in the IUCN Global Ecosystem Typology.",
        "domain": IUCN["Realm"],
        "range": XSD.string,
    },
    "realmType": {
        "label": "realm type",
        "comment": "Type designation of the Realm (e.g. core or transitional).",
        "domain": IUCN["Realm"],
        "range": XSD.string,
    },
    "biomeCode": {
        "label": "biome code",
        "comment": "Code assigned to a Biome in the IUCN Global Ecosystem Typology.",
        "domain": IUCN["Biome"],
        "range": XSD.string,
    },
    "efgCode": {
        "label": "ecosystem functional group code",
        "comment": "Code assigned to an Ecosystem Functional Group in the IUCN Global Ecosystem Typology.",
        "domain": IUCN["EcosystemFunctionalGroup"],
        "range": XSD.string,
    },
    # ── Narrative properties ──
    "keyFeatures": {
        "label": "key features",
        "comment": "Characteristic ecological and functional features of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"],
        "range": RDFS.Literal,
    },
    "distributionSummary": {
        "label": "distribution summary",
        "comment": "Summary description of the distribution of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"],
        "range": RDFS.Literal,
    },
    "ecosystemProperties": {
        "label": "ecosystem properties",
        "comment": "Description of notable structural or functional properties of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"],
        "range": RDFS.Literal,
    },
    "ecologicalDrivers": {
        "label": "ecological drivers",
        "comment": "Description of the main ecological drivers shaping the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"],
        "range": RDFS.Literal,
    },
    "distribution": {
        "label": "distribution",
        "comment": "Detailed description of the geographic or environmental distribution of the ecosystem functional group.",
        "domain": IUCN["EcosystemFunctionalGroup"],
        "range": RDFS.Literal,
        "subproperty_of": SKOS.note,
    },
    # ── Provenance properties ──
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
            "date was used as the top-level dcterms:modified value. Provided "
            "as a flat sibling of dcterms:modified for easy querying."
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


# ── Per-concept emission helper for multilingual rows ─────────────────────────

def emit_translatable(g, uri, lang, fields):
    for pred, val in fields:
        if val:
            g.add((uri, pred, lang_lit(val, lang)))


def get_pref_label(g, uri, lang):
    """Return the prefLabel string for a given language, or None."""
    for o in g.objects(uri, SKOS.prefLabel):
        if isinstance(o, Literal) and o.language == lang:
            return str(o)
    return None


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
                    g.add((uri, SKOS.prefLabel,
                           lang_lit(strip_code_prefix(name, bcode), lang)))
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
    else:
        print("  WARNING: No Biomes sheet found.")

    # ── Level 3: EFGs (multilingual) ─────────────────────────────────────────
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

            # Biome (direct parent) and Realm (transitive ancestor)
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
                    g.add((uri, SKOS.prefLabel,
                           lang_lit(strip_code_prefix(name, code), lang)))
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

        # Sibling EFGs in the same biome → skos:related (symmetric)
        for bcode, uris in efg_by_biome.items():
            for i, a in enumerate(uris):
                for b in uris[i + 1:]:
                    g.add((a, SKOS.related, b))
                    g.add((b, SKOS.related, a))
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

            # altLabel: skip if equal to the English prefLabel
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

            # Per-language web pages (en, es, fr) — these are
            # representations of the concept, not identifiers.
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
    if len(sys.argv) < 2:
        print("Usage: python excel_to_skos_rdf.py <input.xlsx>")
        sys.exit(1)

    import os
    path = sys.argv[1]
    in_dir, in_name = os.path.split(os.path.abspath(path))
    stem = in_name.rsplit(".", 1)[0]

    out_dir = os.path.join(os.path.dirname(in_dir), "output")
    os.makedirs(out_dir, exist_ok=True)

    print(f"\nReading: {path}")
    g = build_graph(path)

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
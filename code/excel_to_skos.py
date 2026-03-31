"""
Excel to SKOS/RDF Converter
============================
Converts a 4-sheet Excel workbook into SKOS/RDF (Turtle + RDF/XML).

Expected sheets:
  - Level 1 realms        : realm, coderealm, typedescription
  - Level 2 biomes        : biome, codenamerealm, codedescription, contributors, version
  - Level 3 ecosystem...  : code, name, ecosystem properties, ecological drivers,
                             distribution, contributors, version
  - Level 3 short desc.   : code, biome code, name, short name, update,
                             short description, key features, distribution summary, url

Reference format (anzsoil.org style):
  - Every concept is typed as both rdfs:Resource AND skos:Concept
  - dcterms:identifier is set as xsd:anyURI
  - String literals are typed as xsd:string (not language-tagged)

Usage:
    pip install pandas openpyxl rdflib
    python excel_to_skos_rdf.py <input.xlsx>

Outputs:
    <input>_skos.ttl   — Turtle (SKOS)
    <input>_skos.rdf   — RDF/XML (SKOS)
"""

import sys
import re
import pandas as pd
from rdflib import Graph, Literal, Namespace, URIRef
from rdflib.namespace import RDF, SKOS, DCTERMS, OWL, XSD, RDFS

# ── Namespaces ────────────────────────────────────────────────────────────────
BASE   = Namespace("https://example.org/iucn/")
IUCN   = Namespace("https://example.org/iucn/ontology#")
SCHEMA = Namespace("https://schema.org/")
FOAF   = Namespace("http://xmlns.com/foaf/0.1/")


# ── Literal helpers matching reference style ──────────────────────────────────

def xsd_str(val):
    """xsd:string typed literal — matches the reference Turtle format."""
    return Literal(str(val), datatype=XSD.string)


def xsd_uri(val):
    """xsd:anyURI typed literal."""
    return Literal(str(val), datatype=XSD.anyURI)


# ── Concept registration ──────────────────────────────────────────────────────

def add_concept(g, uri, extra_type=None):
    """
    Register a URI as rdfs:Resource + skos:Concept (+ optional domain type)
    and set dcterms:identifier as xsd:anyURI.

    Mirrors the reference pattern:
        a  rdfs:Resource , skos:Concept ;
        dcterms:identifier  "..."^^xsd:anyURI ;
    """
    g.add((uri, RDF.type,           RDFS.Resource))
    g.add((uri, RDF.type,           SKOS.Concept))
    if extra_type:
        g.add((uri, RDF.type,       extra_type))
    g.add((uri, DCTERMS.identifier, xsd_uri(str(uri))))


# ── General helpers ───────────────────────────────────────────────────────────

def slugify(text):
    text = str(text).strip()
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"[\s_-]+", "_", text)
    return text.strip("_")


def safe(val):
    if pd.isna(val):
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


# ── Build graph ───────────────────────────────────────────────────────────────

def build_graph(path):
    g = Graph()
    g.bind("rdf",     RDF)
    g.bind("rdfs",    RDFS)
    g.bind("xsd",     XSD)
    g.bind("skos",    SKOS)
    g.bind("dcterms", DCTERMS)
    g.bind("owl",     OWL)
    g.bind("foaf",    FOAF)
    g.bind("iucn",    IUCN)
    g.bind("base",    BASE)
    g.bind("schema",  SCHEMA)

    # Top-level concept scheme
    scheme = BASE["IUCNEcosystemTypology"]
    g.add((scheme, RDF.type,             SKOS.ConceptScheme))
    g.add((scheme, DCTERMS.identifier,   xsd_uri(str(scheme))))
    g.add((scheme, SKOS.prefLabel,       xsd_str("IUCN Global Ecosystem Typology")))
    g.add((scheme, DCTERMS.description,
           xsd_str("A hierarchical classification of Earth's ecosystems.")))

    xl = pd.ExcelFile(path)
    sheet_names = xl.sheet_names
    print(f"Found sheets: {sheet_names}")

    # ── Level 1: Realms ───────────────────────────────────────────────────────
    realm_sheet = (
        next((s for s in sheet_names if "realm" in s.lower() and "level 1" in s.lower()), None)
        or next((s for s in sheet_names if "realm" in s.lower()), None)
    )

    if realm_sheet:
        df = read_sheet(path, realm_sheet)
        print(f"  Level 1 Realms — columns: {list(df.columns)}")
        c_code = col(df, "coderealm", "realm_code", "code")
        c_name = col(df, "realm",     "name",       "realm_name")
        c_desc = col(df, "typedescription", "type_description", "description")

        for _, row in df.iterrows():
            code = safe(row[c_code]) if c_code else None
            name = safe(row[c_name]) if c_name else None
            desc = safe(row[c_desc]) if c_desc else None
            if not code:
                continue

            uri = BASE[f"Realm_{slugify(code)}"]
            add_concept(g, uri, IUCN["Realm"])
            g.add((uri, SKOS.inScheme,       scheme))
            g.add((uri, SKOS.topConceptOf,   scheme))
            g.add((uri, IUCN["realmCode"],   xsd_str(code)))
            if name:
                g.add((uri, SKOS.prefLabel,  xsd_str(name)))
            if desc:
                g.add((uri, SKOS.definition, xsd_str(desc)))
    else:
        print("  WARNING: No Level 1 Realms sheet found.")

    # ── Level 2: Biomes ───────────────────────────────────────────────────────
    biome_sheet = next((s for s in sheet_names if "biome" in s.lower()), None)

    if biome_sheet:
        df = read_sheet(path, biome_sheet)
        print(f"  Level 2 Biomes — columns: {list(df.columns)}")
        c_bcode   = col(df, "biome",         "biome_code", "code")
        c_rcode   = col(df, "codenamerealm", "realm_code", "realm")
        c_name    = col(df, "name",          "biome_name")
        c_desc    = col(df, "codedescription", "description")
        c_contrib = col(df, "contributors",  "contributor")
        c_ver     = col(df, "version")

        for _, row in df.iterrows():
            bcode   = safe(row[c_bcode])   if c_bcode   else None
            rcode   = safe(row[c_rcode])   if c_rcode   else None
            name    = safe(row[c_name])    if c_name    else None
            desc    = safe(row[c_desc])    if c_desc    else None
            contrib = safe(row[c_contrib]) if c_contrib else None
            ver     = safe(row[c_ver])     if c_ver     else None
            if not bcode:
                continue

            uri = BASE[f"Biome_{slugify(bcode)}"]
            add_concept(g, uri, IUCN["Biome"])
            g.add((uri, SKOS.inScheme,           scheme))
            g.add((uri, IUCN["biomeCode"],       xsd_str(bcode)))
            if name:
                g.add((uri, SKOS.prefLabel,      xsd_str(name)))
            if desc:
                g.add((uri, SKOS.definition,     xsd_str(desc)))
            if contrib:
                g.add((uri, DCTERMS.contributor, xsd_str(contrib)))
            if ver:
                g.add((uri, OWL.versionInfo,     xsd_str(ver)))

            if rcode:
                realm_uri = BASE[f"Realm_{slugify(rcode)}"]
                g.add((uri,       SKOS.broader,  realm_uri))
                g.add((realm_uri, SKOS.narrower, uri))
    else:
        print("  WARNING: No Level 2 Biomes sheet found.")

    # ── Level 3: Ecosystem Functional Groups ─────────────────────────────────
    efg_sheet = (
        next((s for s in sheet_names
              if "ecosystem" in s.lower() and "functional" in s.lower()), None)
        or next((s for s in sheet_names
                 if "level 3" in s.lower() and "short" not in s.lower()), None)
    )

    if efg_sheet:
        df = read_sheet(path, efg_sheet)
        print(f"  Level 3 EFG — columns: {list(df.columns)}")
        c_code    = col(df, "code")
        c_name    = col(df, "name")
        c_props   = col(df, "ecosystem_properties", "properties")
        c_drivers = col(df, "ecological_drivers",   "drivers")
        c_dist    = col(df, "distribution")
        c_contrib = col(df, "contributors", "contributor")
        c_ver     = col(df, "version")

        for _, row in df.iterrows():
            code = safe(row[c_code]) if c_code else None
            name = safe(row[c_name]) if c_name else None
            if not code:
                continue

            uri = BASE[f"EFG_{slugify(code)}"]
            add_concept(g, uri, IUCN["EcosystemFunctionalGroup"])
            g.add((uri, SKOS.inScheme,       scheme))
            g.add((uri, IUCN["efgCode"],     xsd_str(code)))
            if name:
                g.add((uri, SKOS.prefLabel,  xsd_str(name)))

            for c_attr, pred in [
                (c_props,   IUCN["ecosystemProperties"]),
                (c_drivers, IUCN["ecologicalDrivers"]),
                (c_dist,    IUCN["distribution"]),
            ]:
                val = safe(row[c_attr]) if c_attr else None
                if val:
                    g.add((uri, pred, xsd_str(val)))

            contrib = safe(row[c_contrib]) if c_contrib else None
            ver     = safe(row[c_ver])     if c_ver     else None
            if contrib:
                g.add((uri, DCTERMS.contributor, xsd_str(contrib)))
            if ver:
                g.add((uri, OWL.versionInfo,     xsd_str(ver)))

            # Infer biome from code prefix (e.g. "T1.1" -> biome "T1")
            parts = code.split(".")
            if len(parts) >= 2:
                biome_code = parts[0]
                biome_uri  = BASE[f"Biome_{slugify(biome_code)}"]
                g.add((uri,       SKOS.broader,  biome_uri))
                g.add((biome_uri, SKOS.narrower, uri))
    else:
        print("  WARNING: No Level 3 EFG sheet found.")

    # ── Level 3: Short Descriptions ──────────────────────────────────────────
    short_sheet = (
        next((s for s in sheet_names
              if "short" in s.lower() and "desc" in s.lower()), None)
        or next((s for s in sheet_names
                 if "level 3" in s.lower() and "short" in s.lower()), None)
    )

    if short_sheet:
        df = read_sheet(path, short_sheet)
        print(f"  Level 3 Short Desc — columns: {list(df.columns)}")
        c_code    = col(df, "code")
        c_bcode   = col(df, "biome_code", "biome")
        c_name    = col(df, "name")
        c_sname   = col(df, "short_name")
        c_update  = col(df, "update")
        c_sdesc   = col(df, "short_description", "short_desc")
        c_feat    = col(df, "key_features")
        c_distsum = col(df, "distribution_summary")
        c_url     = col(df, "url")

        for _, row in df.iterrows():
            code = safe(row[c_code]) if c_code else None
            if not code:
                continue

            uri = BASE[f"EFG_{slugify(code)}"]
            # Register concept if not already present from EFG sheet
            if (uri, RDF.type, SKOS.Concept) not in g:
                add_concept(g, uri, IUCN["EcosystemFunctionalGroup"])
                g.add((uri, SKOS.inScheme,   scheme))
                g.add((uri, IUCN["efgCode"], xsd_str(code)))

            name    = safe(row[c_name])    if c_name    else None
            sname   = safe(row[c_sname])   if c_sname   else None
            sdesc   = safe(row[c_sdesc])   if c_sdesc   else None
            feat    = safe(row[c_feat])    if c_feat    else None
            distsum = safe(row[c_distsum]) if c_distsum else None
            update  = safe(row[c_update])  if c_update  else None
            url     = safe(row[c_url])     if c_url     else None
            bcode   = safe(row[c_bcode])   if c_bcode   else None

            # Only add prefLabel if not already set
            if name and not any(True for _ in g.triples((uri, SKOS.prefLabel, None))):
                g.add((uri, SKOS.prefLabel,              xsd_str(name)))
            if sname:
                g.add((uri, SKOS.altLabel,               xsd_str(sname)))
            if sdesc:
                g.add((uri, SKOS.scopeNote,              xsd_str(sdesc)))
            if feat:
                g.add((uri, IUCN["keyFeatures"],         xsd_str(feat)))
            if distsum:
                g.add((uri, IUCN["distributionSummary"], xsd_str(distsum)))
            if update:
                g.add((uri, DCTERMS.modified,            xsd_str(update)))
            if url:
                g.add((uri, SCHEMA["url"],               URIRef(url)))

            if bcode and not any(True for _ in g.triples((uri, SKOS.broader, None))):
                biome_uri = BASE[f"Biome_{slugify(bcode)}"]
                g.add((uri,       SKOS.broader,  biome_uri))
                g.add((biome_uri, SKOS.narrower, uri))
    else:
        print("  WARNING: No Level 3 Short Descriptions sheet found.")

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
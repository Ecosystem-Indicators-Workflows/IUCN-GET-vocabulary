# Controlled vocabulary for the IUCN Global Ecosystem Typology

This repository includes code and data to create a controlled research vocabulary for the IUCN Global Ecosystem Typology.

## The IUCN Global Ecosystem Typology

The IUCN Global Ecosystem Typology (GET) was developed by Red List of Ecosystems (RLE) Thematic Group of the Commission on Ecosystem Management (CEM) and led to IUCN Publication, dedicated web app (<https://global-ecosystems.org>), and scientific articles.

IUCN is committed to maintaining the Global Ecosystem Typology as the global knowledge base on ecosystems continually grows. 


## Why a controlled vocabulary?

- A controlled vocabulary provides a way to **organize knowledge** for subsequent retrieval.
- **A taxonomy is a type of controlled vocabulary** in which all the terms belong to a single hierarchical structure and have parent/child or broader/narrower relationships to other terms.
- **Users** may only apply terms from a restricted list, only **Editors** can make changes.
- Changes or additions to the list have to meet certain conditions and follow **review processes**.
- **Designing the metadata elements** for a digital asset collection and designing the controlled vocabulary are integrated processes.

## Making a FAIR vocabularies

Following recommendations from Cox et al. (2021) 

### Rule 1. ✅ Determine the governance arrangements and custodian of the legacy vocabulary

The ***Global Ecosystem Typology Scientific Committee (GET-SC)*** was established in 2024 to advise on scientific and technical aspects of the maintenance and update of the GET to ensure high standards of scientific excellence and its fitness for purpose across diverse applications.

We will coordinate with the [GET-SC and the GET Data hub](https://red-list-ecosystem.github.io/GET-data-hub/governance.html)

### Rule 2. ✅ Verify that the legacy-vocabulary license allows repurposing, and agree on the license for the FAIR vocabulary

All content is based on Keith et al. (2022) and reuse of the text for the vocabulary complies with relevant licenses and [terms of use](https://global-ecosystems.org/page/terms).

This FAIR vocabulary in the [output](/output/) folder is licensed under CC-BY license.

### Rule 3. ✅ Check term and definition completeness and consistency in the legacy vocabulary

All terms and definitions are complete and consistent with the [updated on-line version of the GET](https://global-ecosystems.org/).

### Rule 4. ✅ Establish a traceable maintenance-environment for the FAIR vocabulary content

This [repository](https://github.com/Ecosystem-Indicators-Workflows/IUCN-GET-vocabulary) was create to keep a version control and issue tracking of the changes to the vocabulary.


### Rule 5. ✅ Assign a unique and persistent identifier to (a) the vocabulary and (b) each term in the vocabulary

We registered the permanent ID [w3id.org/iucn-get](https://w3id.org/iucn-get/) for this vocabulary.

### Rule 6. ✅ Create machine readable representations of the vocabulary terms

Scripts to translate input tables into machine readable representations are located in folder [code](/code/).

The code is shared with a [MIT License](/code/LICENSE).

### Rule 7. 🚧👷🏽 Add vocabulary metadata

Work in progress

### Rule 8. 🚧👷🏽 Register the vocabulary 

We loaded / registered the encoded content in the following vocabulary service: [*Research Vocabularies Australia*](https://vocabs.ardc.edu.au/) 

Currently operating as a demo version: [demo.vocabs.ardc.edu.au](https://demo.vocabs.ardc.edu.au/viewById/1167)


### Rule 9. 🚧👷🏽 Make the vocabulary accessible for humans and machines

Work in progress.

### Rule 10. ⏳ Implement a process for publishing revisions of the FAIR vocabulary

To do.

## The vocabulary in action

### Within the Ecosystem Indicator Workflows project

- [ecoSynthesis](https://github.com/Ecosystem-Indicators-Workflows/ecoSynthesis) is an R package for synthesis of cross-referencing information between two classification systems. See [documentation](https://ecosystem-indicators-workflows.github.io/ecoSynthesis/).
- Used in several examples of [Workflows for cross-referencing Australian Ecosystems](https://github.com/search?q=topic%3Across-referencing+org%3AEcosystem-Indicators-Workflows&type=Repositories)
- [EFG reference data store](https://github.com/Ecosystem-Indicators-Workflows/EFG-reference-data-store) provides input data for EcoID tool (under development).

### Outside the EIW project
- [iucn-get-data](https://github.com/RLE-Assessment/iucn-get-data) provides Python access to the upper three levels of the IUCN Global Ecosystem Typology (GET) 2.0. See [documentation](https://rle-assessment.github.io/iucn-get-data/).
- IUCN Red List of Ecosystems assessments database ([SERIS project](https://findanexpert.unimelb.edu.au/project/26-2350-spatial-ecosystem-risk-intelligence-system-(seris)), under development)



## References

- Keith, D. A., J. R. Ferrer-Paris, E. Nicholson, M. Bishop, B. A, Polidoro, E. Ramirez-Llodra, M. G. Tozer, J. L. Nel, R. Mac Nally, E. J. Gregr, K. E. Watermeyer, F. Essl, D. Faber-Langendoen, J. Franklin, C. E. R. Lehmann, A. Etter, D. J. Roux, J. S. Stark, J. A. Rowland, N. A. Brummitt, U. C. Fernandez-Arcaya, I. M. Suthers, S. K. Wiser, I. Donohue, L. J. Jackson, R. T. Pennington, N. Pettorelli, A. Andrade, A. Lindgaard, T. Tahvanainen, A. Terauds, M. A. Chadwick, N. J. Murray, J. Moat, P. Pliscoff, I. Zager, and R. T. Kingsford (2022) A function-based typology for Earth’s ecosystems Nature 610, 513–518. DOI:10.1038/s41586-022-05318-4.
- Cox SJtD, Gonzalez-Beltran AN, Magagna B, Marinescu M-C (2021) Ten simple rules for making a vocabulary FAIR. PLoS Comput Biol 17(6): e1009041. https://doi.org/10.1371/journal.pcbi.1009041

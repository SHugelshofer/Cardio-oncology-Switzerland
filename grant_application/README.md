# Grant Application Materials

This directory contains materials for grant applications related to the CVRF-CARDIO-PROTECT study (Cardiovascular Risk Factors, Guideline Adherence, and Cardiotoxicity in Cardio-Oncology, with a Focus on Diabetes and SGLT2 Inhibitors).

## Contents

### Project Documentation (Markdown & Word formats)

- **resume_projet_SGLT2_diabetes.md** - Comprehensive project summary with dedicated focus on diabetes, glycaemic control, and SGLT2 inhibitor cardioprotection (English, for Boehringer Ingelheim submission)
- **resume_projet.md / .docx** - General project summary (French, for Fondation Emma Muschamp)
- **budget_detaille.md / .docx** - Detailed budget breakdown for both funding sources (CHF 50,000 Fondation Emma Muschamp + CHF 20,000 Boehringer Ingelheim; total project CHF 147,000)
- **budget_adapte_50k_12mois.md** - Budget adapté pour CHF 50,000 sur 12 mois (Fondation Emma Muschamp uniquement)
- **references_epidemiologie.md / .docx** - Scientific references and sources for cardiovascular risk factor prevalence statistics
- **demographics_age_cancer.md / .docx** - Detailed age and cancer type distribution from cited studies
- **sources_statistiques_tableau.md / .docx** - Quick-reference table mapping statistics to sources

### Support and Request Letters

- **lettre_Boehringer_Ingelheim.md** - Grant request letter to Boehringer Ingelheim (CHF 20,000) — focused on diabetes and SGLT2 inhibitor cardioprotection component (English)
- **lettre_soutien_Prof_Muller.md / .docx** - Professional support letter from Prof. Olivier Müller (Chef du Service de Cardiologie, CHUV) for the CHF 50,000 grant application to Fondation Emma Muschamp
- **lettre_Novartis_Dr_Monsutti.md / .docx** - Letter to Novartis (Dr Alice Monsutti)

### Converting to Word Format

All markdown files can be converted to Microsoft Word (.docx) format using the conversion script:

```bash
# Convert all documents at once
python3 ../convert_to_word.py --all

# Convert specific documents
python3 ../convert_to_word.py grant_application/resume_projet.md grant_application/budget_detaille.md

# Convert to a specific output directory
python3 ../convert_to_word.py --all --output-dir ./word_documents
```

**Note:** Word documents (.docx) are automatically generated and are ideal for:
- Submission to grant committees (Fondation Emma Muschamp)
- Printing on official CHUV letterhead
- Editing and formatting by committee members
- Distribution to collaborators

## Usage Instructions

### For the Project Summary

The project summary (`resume_projet.md`) provides:
- Complete background and rationale for the study
- **Evidence-based cardiovascular risk factor prevalence rates** (with citations)
- Detailed methodology and statistical analysis plan
- Expected outcomes and impact
- Budget breakdown and timeline

**Important Note on Statistics:** The prevalence rates cited (diabète 20-30%, hypertension 35-50%, dyslipidémie 30-40%, obésité 25-35%, maladie cardiovasculaire établie 15-25%) are based on peer-reviewed literature. See `references_epidemiologie.md` for complete citations including:
- Koene RJ et al., Circulation 2016
- Herrmann J et al., Circulation Research 2022
- Armenian SH et al., JACC: CardioOncology 2023 (CARDIOTOX: N=10,052, median age 54)
- Lyon AR et al., ESC Heart Failure 2022 (EC3: N=8,421, median age 58)
- Sturgeon KM et al., European Heart Journal 2019 (SEER: N=3.2M, median age 64)
- And other major cardio-oncology studies

**Age and Cancer Distribution:** See `demographics_age_cancer.md` for detailed breakdowns:
- CARDIOTOX Registry: Median age 54 years; Breast cancer 32%, Lymphomas 19%
- EC3 European Study: Median age 58 years; Breast cancer 38%, Lymphomas 16%
- SEER Registry: Median age 64 years; Prostate 23%, Breast 22%, Lung 13%

### For the Boehringer Ingelheim Grant Request Letter

The grant request letter (`lettre_Boehringer_Ingelheim.md`) is written in English and:
- Targets a specific CHF 20,000 grant from Boehringer Ingelheim
- Focuses exclusively on the diabetes and SGLT2 inhibitor cardioprotection component
- Should be printed on official CHUV letterhead and signed by Dr. Hugelshofer and Prof. Müller
- Describes the complementarity with the Emma Muschamp Foundation grant

### For the SGLT2/Diabetes Focused Project Summary

The `resume_projet_SGLT2_diabetes.md` document:
- Is written in English for international/pharmaceutical audience (Boehringer Ingelheim)
- Provides detailed background on SGLT2 inhibitor cardioprotective mechanisms
- Describes the HFA-PEFF score assessment embedded in the study design
- Includes specific SGLT2 inhibitor-focused secondary endpoints and analyses

### For the Support Letter (Fondation Emma Muschamp)

The support letter (`lettre_soutien_Prof_Muller.md`) should be:

1. **Printed on official CHUV letterhead** - This is essential for authenticity
2. **Reviewed and signed** by Prof. Olivier Müller
3. **Dated** appropriately before submission
4. **Submitted** as part of the grant application package to Fondation Emma Muschamp

The letter is written in professional French and emphasizes:
- Scientific merit and innovation of the CVRF-CARDIO-PROTECT project
- Dr. Sarah Hugelshofer's qualifications and expertise
- Institutional support from CHUV (infrastructure, REDCap database, statistician, office space)
- Expected impact and publication strategy
- Alignment with CHUV's strategic priorities in cardio-oncology

## Project Information

**Project Name:** CVRF-CARDIO-PROTECT  
**Full Title:** Cardiovascular Risk Factors, Guideline Adherence, and Cardiotoxicity in Cardio-Oncology with a Focus on Diabetes and SGLT2 Inhibitors  
**Principal Investigator:** Dr. Sarah Hugelshofer  
**Institution:** CHUV - Centre Hospitalier Universitaire Vaudois  
**Department:** Service de Cardiologie  
**Requested Amount:** CHF 70,000 (CHF 50,000 Fondation Emma Muschamp + CHF 20,000 Boehringer Ingelheim)  
**Total Project Budget:** CHF 147,000  
**Funding Sources:** Fondation Emma Muschamp, Boehringer Ingelheim, other competitive sources, CHUV in-kind

## Contact

For questions about this grant application or the CVRF-CARDIO-PROTECT study, please contact:

**Dr. Sarah Hugelshofer**  
Service de Cardiologie  
CHUV  
Email: [contact information]

**Prof. Olivier Müller**  
Chef du Service de Cardiologie  
CHUV  
Email: olivier.muller@chuv.ch  
Tel: +41 21 314 11 11

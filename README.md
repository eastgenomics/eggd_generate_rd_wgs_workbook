<!-- dx-header -->

# egg_generate_rd_wgs_workbook (DNAnexus Platform App)

## What does this app do?

Generates an Excel workbook from a Genomics England rare disease case JSON.


## What data are required for this app to run?

**Packages**
* Python packages (specified in requirements.txt)

**Inputs**
Required
* `json`: a GEL RD WGS JSON with data to be put into a workbook
* `refseq_tsv`: a RefSeq TSV with Ensembl protein and transcript IDs
* `mane_file`: a MANE csv with ensembl -> refseq transcript ID conversions for MANE transcripts only.
* `config`: config file with DNAnexus file IDs for HPO obo files, and thresholds for de novo quality score for SNVs and CNVs

Optional
* `acmg`: number of SNV ACMG interpretation sheets to add to the workbook
* `cnv`: number of CNV interpretation sheets to add to the workbook
* `epic_clarity`: .xlsx file of Epic Clarity export, will be used to add sample IDs to the workbook

## What variants are displayed in the workbook?
GEL Tiering sheet
* TIER1 SNVs
* TIER2 SNVs
* TIERA/TIER1 CNVs
* TIER1 STRs
Extended analysis sheet
* Exomiser Top 3 ranked variants
    * only displayed if:
        * score >= 0.75
        * variant is not mitochondrial and untiered.
        * variant has not already been reported in GEL tiering sheet.
* De novo variants:
    * all variants with a de novo score above the threshold, regardless of tier.

## What does this app output?

This app outputs an Excel workbook
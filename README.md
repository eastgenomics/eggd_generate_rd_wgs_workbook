<!-- dx-header -->

# egg_generate_rd_wgs_workbook (DNAnexus Platform App)

## What does this app do?

Generates an Excel workbook from a Genomics England rare disease case JSON.


## What data are required for this app to run?

**Packages**
* Python packages (specified in requirements.txt)

**Inputs**
Required
* `json`: a GEL JSON with data to be put into a workbook
* `refseq_tsv`: a RefSeq TSV with Ensembl protein and transcript IDs
* `mane`: a MANE csv with ensembl -> refseq transcript ID conversions for MANE transcripts only.
* `obo_path` OR `obo_files`: the path to a directory with .obo files stored locally, or a DNAnexus array of obo files

Optional
* `acmg`: number of SNV ACMG interpretation sheets to add to the workbook
* `cnv`: number of CNV interpretation sheets to add to the workbook
* `output`: file name for workbook, if not set with use JSON input file name
* `out_dir`: path to output directory.

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
* DeNovo variants:
    * all variants with a de novo score above the threshold, regardless of tier.

## What does this app output?

This app outputs an Excel workbook
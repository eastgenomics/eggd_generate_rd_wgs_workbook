import dxpy
import networkx
import json
import obonet
import openpyxl.drawing
import openpyxl.drawing.image
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, DEFAULT_FONT, Font, Side
from openpyxl.styles.fills import PatternFill
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
import os
from pathlib import Path
import re

# openpyxl style settings
THIN = Side(border_style="thin", color="000000")
MEDIUM = Side(border_style="medium", color="000001")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

DEFAULT_FONT.name = 'Calibri'

# row and col counts that are to be unlocked next to
# populated table in all sheets if it is dias pipeline
# required for 'lock_sheet' function
ROW_TO_UNLOCK = 500
COL_TO_UNLOCK = 200

class excel():
    '''
    Functions to generate excel workbook of variants
    '''
    def __init__(self, args) -> None:
        self.args = args
        self.wgs_data = None
        self.writer = None
        self.workbook = None
        self.summary = None
        self.proband = None
        self.mother = None
        self.father = None
        self.bold_content = None
        self.summary_content = None
        self.mane = None
        self.refseq_tsv = None
        self.gel_index = None
        self.ex_index = None
        self.var_df = None
        self.column_list = [
                "Gene",
                "Chr",
                "Pos",
                "End",
                "Length",
                "Ref",
                "Alt",
                'STR1',
                'STR2',
                "Repeat",
                "Copy Number",
                "Type",
                "HGVSc",
                "HGVSp",
                "Priority",
                "Inheritance",
                "Segregation pattern",
                "Inheritance mode",
                "Zygosity",
                "Depth",
                "AF Max",
                "Penetrance filter",
                "Comment",
                "Checker comment"
            ]

    def generate(self) -> None:
        """
        Calls all methods in excel() to generate output file.
        """
        # Initiate files and workbook to write in
        self.open_files()
        self.writer = pd.ExcelWriter(self.args.output, engine='openpyxl')
        self.workbook = self.writer.book
        print(f"Writing to {self.args.output}...")
        # Write in workbook
        self.summary_page()
        self.find_interpretation_service()
        self.create_gel_tiering_variant_page()
        self.create_additional_analysis_page()
        self.str_image_page()
        self.writer.close()
        if self.args.acmg:
            for i in range(1, self.args.acmg+1):
                self.write_reporting_template(i)
        if self.args.cnv:
            for i in range(1, self.args.cnv+1):
                self.write_cnv_reporting_template(i)
        
        # self.workbook.save(self.args.output)
        # if self.args.acmg and self.args.lock_sheet:
        #     self.protect_rename_sheets()
        # if self.args.acmg:
        #     self.drop_down()
        self.workbook.save(self.args.output)
        self.drop_down()

        print('Done!')

    def summary_page(self):
        '''
        Add summary page. Create a page in the workbook to populate with
        details about the case and variants for interpretation.
        '''
        summary_sheet = self.workbook.create_sheet("Summary")

        self.bold_content = {
            (1, 1): "Family ID",
            (6, 1): "Proband",
            (7, 1): "Mother",
            (8, 1): "Father",
            (5, 2): "ID",
            (5, 3): "GM/SP number",
            (5, 4): "Affected?",
            (5, 5): "HPO",
            (2, 1): "Clinical indication",
            (1, 8): "Flags",
            (2, 8): "Test code",
            (3, 1): "Penetrance",
            (30, 1): "Overall result",
            (31, 1): "Confirmation",
            (33, 2): "Name",
            (33, 3): "Date",
            (34, 1): "Primary analysis",
            (35, 1): "Data check",
            (10, 1): "LP number",
            (12, 1): "Panels",
            (13, 1): "Panel ID",
            (13, 2): "Indication",
            (13, 3): "Version",
            (13, 4): "Panel name",
            (19, 1): "Tiered variants",
            (19, 2): "In this case",
            (19, 3): "To be reported",
            (25, 1): "Extended analysis",
            (25, 2): "In this case",
            (25, 3): "To be reported",
            (20, 1): "SNV Tier 1",
            (21, 1): "SNV Tier 2",
            (22, 1): "CNV Tier 1",
            (23, 1): "STRs",
            (23, 1): "STRs",
            (26, 1): "Exomiser top 3 (score ≥ 0.75)",
            (27, 1): "De novo",
        }

        self.summary_content = {
            (1, 9): str(self.wgs_data["interpretation_flag"]),
            (2, 2): self.wgs_data['referral']["clinical_indication_full_name"],
            (1, 2): self.wgs_data["family_id"],
            (2, 9): self.wgs_data['referral']["clinical_indication_code"]
        }

        # Add panel data, penetrance data and data about family members
        self.get_panels()
        self.get_penetrance()
        self.person_data()

        # write summary content and titles to page
        for key, val in self.bold_content.items():
            summary_sheet.cell(key[0], key[1]).value = val
            summary_sheet.cell(key[0], key[1]).font = Font(
                bold=True, name=DEFAULT_FONT.name
            )
        for key, val in self.summary_content.items():
            summary_sheet.cell(key[0], key[1]).value = val

        # Set column widths
        summary_sheet.column_dimensions["A"].width = 20
        summary_sheet.column_dimensions["B"].width = 16
        summary_sheet.column_dimensions["C"].width = 14
        summary_sheet.column_dimensions["D"].width = 14

        row_ranges = {
            'horizontal': [
                'A33:C33', 'A34:C34', 'A36:C36'
            ],
            'vertical':[
                'A33:A35', 'B33:B35', 'C33:C35', 'D33:D35'
            ]
        }

        self.borders(row_ranges, summary_sheet)

    def hpo_version(self, version):
        '''
        Select which version of HPO to use based on which version was used by
        GEL when the JSON was made.
        Works with input obo_files (DNAnexus array of obo files) and input 
        obo_path (local path to directory containing obo files)
        Inputs:
            version (str): version of HPO used in JSON
        Outputs:
            obo (str): Path to HPO obo file for the version of HPO used in the
            JSON
        '''
        if self.args.obo_files:
            if version == "v2019_02_12":
                obo="/home/dnanexus/obo_files/hpo_v20190212.obo"
            elif version == "releases/2018-10-09":
                obo="/home/dnanexus/obo_files/hpo_v20181009.obo"
            else:
                raise RuntimeError(
                    f"GEL version of HPO ontology {version} does not match "
                    "provided obo file(s)"
                )

        elif self.args.obo_path:
            if self.args.obo_path.endswith('/'):
                self.args.obo_path = self.args.obo_path[:-1]

            if version == "v2019_02_12":
                obo = f"{self.args.obo_path}/hpo_v20190212.obo"
            elif version == "releases/2018-10-09":
                obo = f"{self.args.obo_path}/hpo_v20181009.obo"
            else:
                raise RuntimeError(
                    f"GEL version of HPO ontology {version} does not match "
                    "provided obo file(s)"
                )
        else:
            raise RuntimeError(
                "Neither obo_path or obo_files input provided. Process "
                "requires HPO obo files to complete."
            )

        return obo

    def get_hpo_terms(self, member):
        '''
        Use obo hpo term ontology (.obo) file to convert HPO IDs to names
        Inputs:
            member (dict): family member dictionary extracted from GEL JSON
        Outputs:
            hpo_names (list): list of HPO names corresponding to HPO terms for
            that family member
        '''
        hpo_terms = []
        hpo_names = []
        if member["hpoTermList"]:
            # get HPO version from JSON and select that obo file
            obo = self.hpo_version(member["hpoTermList"][0]['hpoBuildNumber'])

            graph = obonet.read_obo(obo)
            # Read in HPO IDs from JSON
            for i in member["hpoTermList"]:
                hpo_terms.append(i["term"])
            # Convert term to name using obo file
            for term in hpo_terms:
                hpo_dict = graph.nodes[term]
                hpo_name = hpo_dict['name']
                hpo_names.append(hpo_name)

            hpo_names = '; '.join(hpo_names)

        else:
            hpo_names = None

        return hpo_names

    def person_data(self):
        '''
        Find data for participants and add to summary sheet.
        This function will find the proband and add their affected status, HPO
        term names, participant ID, sample ID to the summary sheet
        It will also do this for any relatives in the JSON
        '''
        num_participants = 0
        for member in self.wgs_data['referral']['referral_data']["pedigree"]["members"]:
            if member["isProband"] == True:
                num_participants += 1
                terms_list = self.get_hpo_terms(member)
                self.summary_content[(6, 5)] = terms_list
                self.summary_content[(6, 4)] = member["affectionStatus"]
                self.summary_content[(6, 2)] = member["participantId"]
                self.proband = member["participantId"]
                self.summary_content[(10, 2)] = member["samples"][0]["sampleId"]

            elif member["additionalInformation"]["relation_to_proband"] == "Mother":
                num_participants += 1
                terms_list = self.get_hpo_terms(member)
                self.summary_content[(7, 5)] = terms_list
                self.summary_content[(7, 4)] = member["affectionStatus"]
                self.summary_content[(7, 2)] = member["participantId"]
                self.mother = member["participantId"]

            elif member["additionalInformation"]["relation_to_proband"] == "Father":
                num_participants += 1
                terms_list = self.get_hpo_terms(member)
                self.summary_content[(8, 5)] = terms_list
                self.summary_content[(8, 4)] = member["affectionStatus"]
                self.summary_content[(8, 2)] = member["participantId"]
                self.father = member["participantId"]
            else:
                num_participants += 1
                self.summary_content[(9, 1)] = member["additionalInformation"][
                    "relation_to_proband"
                    ]
                terms_list = self.get_hpo_terms(member)
                self.summary_content[(9, 5)] = terms_list
                self.summary_content[(9, 4)] = member["affectionStatus"]
                self.summary_content[(9, 2)] = member["participantId"]
                self.father = member["participantId"]

        if num_participants != len(
            self.wgs_data['referral']['referral_data']["pedigree"]["members"]
            ):
            raise RuntimeError(
                "Number of participants found does not equal number of family"
                " members in JSON."
            )

    def get_panels(self):
        '''
        Function to get panels from JSON and add to summary content dict to
        then add to summary content sheet
        '''
        panels = []
        for test in self.wgs_data['referral']['referral_data'][
            'referralTests'
            ]:
            for panel in test['analysisPanels']:
                panel_list = [
                    panel['panelId'],
                    panel['panelName'],
                    panel['panelVersion'],
                    panel['specificDisease']
                ]
                panels.append(panel_list)

        end = 1 + int(len(panels))
        for i in range(1, end):
            self.summary_content[(i+13, 1)] = panels[i-1][0]
            self.summary_content[(i+13, 2)] = panels[i-1][3]
            self.summary_content[(i+13, 3)] = panels[i-1][2]
            self.summary_content[(i+13, 4)] = panels[i-1][1]

    def get_penetrance(self):
        '''
        Get penetrance info from JSON and add to summary content dict to
        then add to summary content sheet
        '''
        for penetrance in self.wgs_data['referral']['referral_data'][
            'pedigree'
            ]['diseasePenetrances']:
            if penetrance['specificDisease'] == self.summary_content[(2,2)]:
                disease_penetrance = penetrance["penetrance"]

        self.summary_content[(3, 2)] = disease_penetrance

    def str_image_page(self):
        '''
        STR table is useful for interpretation, so will be included on a sheet
        so it can be referred to during interpretation.
        '''
        str_sheet = self.workbook.create_sheet("STR guidelines")
        script_dir = os.path.dirname(__file__)
        img_folder = "images/str_table.png"
        img_path = os.path.join(script_dir, img_folder)
        img = openpyxl.drawing.image.Image(img_path)
        img.anchor = 'B4'
        str_sheet.add_image(img)
        str_sheet['B2'] = (
            "From CU-WG-REF-40 Guidelines for Rare Disease Whole Genome "
            "Sequencing & Next Generation Sequencing Panel Interpretation & "
            "Reporting"
        )

    def read_dx_file(self, file):
        '''
        read a dx file
        '''
        print(f"Reading from {file}")

        if isinstance(file, dict):
            # provided as {'$dnanexus_link': '[project-xxx:]file-xxx'}
            file = file.get('$dnanexus_link')

        if re.match(r'^file-[\d\w]+$', file):
            # just file-xxx provided => find a project context to use
            file_details = self.get_file_project_context(file)
            project = file_details.get('project')
            file_id = file_details.get('id')
        elif re.match(r'^project-[\d\w]+:file-[\d\w]+', file):
            # nicely provided as project-xxx:file-xxx
            project, file_id = file.split(':')
        else:
            # who knows what's happened, not for me to deal with
            raise RuntimeError(
                f"DXFile not in an expected format: {file}"
            )
        
        return dxpy.DXFile(
            project=project, dxid=file_id).read().rstrip('\n').split('\n')

    def get_file_project_context(self, file) -> dxpy.DXObject:
        """
        Get project ID for a given file ID, used where only file ID is
        provided as DXFile().read() requires both, will ensure that
        only a live version of a project context is returned.

        Parameters
        ----------
        file : str
            file ID of file to search

        Returns
        -------
        DXObject
            DXObject file handler object

        Raises
        ------
        AssertionError
            Raised if no live copies of the given file could be found
        """
        print(f"Searching all projects for: {file}")

        # find projects where file exists and get DXFile objects for
        # each to check archivalState, list_projects() returns dict
        # where key is the project ID and value is permission level
        projects = dxpy.DXFile(dxid=file).list_projects()
        print(f"Found file in {len(projects)} project(s)")

        files = [
            dxpy.DXFile(dxid=file, project=id).describe()
            for id in projects.keys()
        ]

        # filter out any archived files or those resolving
        # to the current job container context
        files = [
            x for x in files
            if x['archivalState'] == 'live'
            and not re.match(r"^container-[\d\w]+$", x['project'])
        ]
        assert files, f"No live files could be found for the ID: {file}"

        print(
            f"Found {file} in {len(files)} projects, "
            f"using {files[0]['project']} as project context"
        )

        return files[0]

    def open_files(self):
        '''
        Open input files and read into variables.
        '''

        if 'dnanexus' in self.args.json:
            json_dict = json.loads(self.args.json)
            self.wgs_data = self.read_dx_file(json_dict)
            self.wgs_data = json.loads(self.wgs_data[0])
            # set output prefix to family id, otherwise the file is named the
            # dx file id for the JSON (not ideal)
            self.args.output = self.wgs_data["family_id"] + ".xlsx"
        else:
            with open(self.args.json) as f:
                self.wgs_data = json.load(f)

        if 'dnanexus' in self.args.mane:
            mane_dict = json.loads(self.args.mane)
            self.mane = self.read_dx_file(mane_dict)
        else:
            with open(self.args.mane) as f:
                self.mane = f.readlines()

        if 'dnanexus' in self.args.refseq_tsv:
            refseq_tsv_dict = json.loads(self.args.refseq_tsv)
            self.refseq_tsv = self.read_dx_file(refseq_tsv_dict)
        else:
            with open(self.args.refseq_tsv) as refseq_tsv:
                self.refseq_tsv = refseq_tsv.readlines()

    def create_gel_tiering_variant_page(self):
        '''
        Take variants from GEL tiering JSON and format into sheet in Excel workbook.
        '''
        variant_list = []

        for variant in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"]["variants"]:
            proband_index = self.check_if_proband(variant["variantCalls"])

            for event in variant["reportEvents"]:                
                if event["tier"] in ["TIER1", "TIER2"]:
                    event_index = variant["reportEvents"].index(event)
                    var_dict = self.get_snv_info(
                        variant, proband_index, event_index
                        )
                    c_dot, p_dot = self.get_hgvs_gel(variant)
                    var_dict["HGVSc"] = c_dot
                    var_dict["HGVSp"] = p_dot
                    variant_list.append(var_dict)


        for str_variant in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"][
                "shortTandemRepeats"
            ]:
            proband_index = self.check_if_proband(
                str_variant["variantCalls"]
                )

            for event in str_variant["reportEvents"]:                
                if event["tier"] in ["TIER1", "TIER2"]:
                    var_dict = self.get_str_info(
                        str_variant, proband_index 
                    )
                    variant_list.append(var_dict)
        
        for sv in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"]["structuralVariants"]:
           
            for event in sv["reportEvents"]:
                event_index = sv["reportEvents"].index(event)
                # CNVs can be reported as Tier 1 or Tier A, GEL updated the
                # nomenclature in 2024
                if sv["reportEvents"][event_index]["tier"] in [
                    "TIER1", "TIERA"
                    ]:
                    var_dict = self.get_cnv_info(sv, event_index)
                    variant_list.append(var_dict)

        self.var_df = pd.DataFrame(variant_list)
        self.var_df = self.var_df.drop_duplicates()
        self.var_df['Depth'] = self.var_df['Depth'].astype(object)
        
        # if df is not empty
        if not self.var_df.empty:
            # if both Tier 1 and Tier 2, keep Tier 1 entry only
            self.var_df = self.var_df.sort_values(
                by=['Priority']
                ).drop_duplicates(subset=['Chr','Pos','Ref', 'Alt', 'End'])
            # Sort by Priority and then Gene symbol
            self.var_df = self.var_df.sort_values(['Priority', 'Gene'])
            self.var_df.to_excel(
                self.writer, sheet_name="Variants", index=False
            )

        # Set column widths
        self.resize_variant_columns(self.workbook["Variants"])

        # Add counts to summary sheet
        summary_sheet = self.workbook["Summary"]
        count_dict = {
            'B20': "TIER1_SNV",
            'B21': "TIER2_SNV",
            'B22': "TIER1_CNV",
            'B23': "TIER1_STR",
        }
        for key, val in count_dict.items():
            if val in self.var_df.Priority.values:
                summary_sheet[key] = self.var_df['Priority'].value_counts()[val]
            else:
                summary_sheet[key] = 0

    def find_interpretation_service(self):
        '''
        The JSON contains two interpretation services: GEL tiering and Exomiser
        This function finds the index for each so these can be referred to
        correctly and sets self.gel_index to the index for the GEL tiering and
        self.ex_index to to the index for the Exomiser interpretation.
        '''
        for interpretation in self.wgs_data["interpretedGenomes"]:
            if interpretation['interpretedGenomeData'][
                'interpretationService'
                ] == 'genomics_england_tiering':
                self.gel_index = self.wgs_data[
                    "interpretedGenomes"
                    ].index(interpretation)
            elif interpretation['interpretedGenomeData'][
                'interpretationService'
                ] == 'Exomiser':
                self.ex_index = self.wgs_data[
                    "interpretedGenomes"
                    ].index(interpretation)
            else:
                raise RuntimeError(
                    "Interpretation services in JSON not recognised as "
                    "'genomics_england_tiering' or 'Exomiser'"
                )

    def add_columns_to_dict(self):
        '''
        Function to add columns to variant pages.
        All columns are empty strings, so they can be overwritten by values if
        needed, but left blank if not.
        '''
        variant_dict = {}
        for col in self.column_list:
            variant_dict[col] = ""

        return variant_dict

    def convert_ensembl_to_refseq(self, ensembl):
        '''
        Search MANE file for query ensembl transcript ID. If match found,
        output the equivalent refseq ID, else output None
        Inputs:
            ensembl (str): Ensembl transcript ID from JSON
        Outputs:
            refseq (str): Matched MANE refseq ID to Ensembl transcript ID, or
            None if no match found.
        '''
        refseq = None

        for line in self.mane:
            if ensembl in line:
                refseq = line.split(',')[3].strip("\"")
                break
        return refseq

    def get_af_max(self, variant):
        '''
        Get AF for population with highest allele frequency in the JSON
        Inputs:
            variant (dict): dict describe a single variant from JSON
        Outputs:
            highest_af (int): frequency of the variant in the population with
            the highest allele frequency of the variant, or 0 if the variant
            is not seen in any populations in the JSON
        '''
        # TODO change so this includes all populations!
        if variant['variantAttributes']['alleleFrequencies'] is not None:
            filtered_AFs = []
            for allele_freq in variant['variantAttributes']['alleleFrequencies']:
                if allele_freq['study'] in [
                    'GNOMAD_EXOMES',
                    "GNOMAD_GENOMES"
                    ] and allele_freq['population'] not in [
                        "OTH",
                        "ASJ",
                        "FIN",
                        "MALE",
                        "FEMALE"
                    ]:
                    # filter out Ashkenazi Jewish, Finnish and Other populations
                    # these are filtered out by gnomad and congenica when working
                    # out AF max for popmax/grpmax
                    # Also filter out Male/Female populations
                    filtered_AFs.append(allele_freq)
                if allele_freq['study'] in [
                    "GEL_aggCOVID_V4.2-20220117",
                    "CNV_AF",
                    "CNV_AUC"
                    ]:
                    filtered_AFs.append(allele_freq)
            
            highest_af = 0
            for af in filtered_AFs:
                if af['alternateFrequency'] > highest_af:
                    highest_af = af['alternateFrequency']
        else:
            highest_af = 0
        return int(highest_af)

    def get_gene_symbol(self, variant):
        '''
        Get gene symbol from variant record
        '''
        gene_list = []
        for entry in variant['reportEvents'][0]['genomicEntities']:
            if entry['type'] == 'gene':
                gene_list.append(entry['geneSymbol'])
        uniq_genes = list(set(gene_list))
        if len(uniq_genes) == 1:
            gene_symbol = uniq_genes[0]
        else:
            gene_symbol = str(uniq_genes).strip('[').strip(']')
        return gene_symbol

    def get_ensp(self, enst):
        '''
        from input ensembl transcript ID, get ensembl protein ID from the
        refseq_tsv
        Inputs:
            enst (str): Ensembl transcript ID from JSON
        Outputs:
            ensp (str): Ensembl protein ID equivalent to the input transcript
            ID
        '''
        ensp = None
        for line in self.refseq_tsv:
            if enst in line:
                ensp = [x for x in line.split() if x.startswith('ENSP')]
                ensp = ensp[0]
                break
        
        return ensp

    def get_hgvs_exomiser(self, variant):
        '''
        Exomiser variants have HGVS nomenclature for p dot and c dot provided
        in one field in the JSON in the following format:
        gene_symbol:ensembl_transcript_id:c_dot:p_dot
        This function extracts the cdot (with MANE refseq equivalent to 
        ensembl transcript ID if found) and pdot (with ensembl protein ID)
        Inputs:
            variant: (dict) dict describing single variant from JSON
        Outputs:
            hgvs_c: (str) HGVS c dot (transcript) nomenclature for the variant.
            annotated against refseq transcript ID if one exists, and ensembl
            transcript ID if there is no matched equivalent
            hgvs_p: (str) HGVS p dot (protein) nomenclature for the variant.
            this is annotated against the ensembl protein ID.
        '''
        hgvs_c = None
        hgvs_p = None
        hgvs_source = variant['variantAttributes'][
            'additionalTextualVariantAnnotations'
            ]['hgvs']
        refseq = self.convert_ensembl_to_refseq(hgvs_source.split(':')[1])
        if refseq is not None:
            hgvs_c = refseq + ":" + hgvs_source.split(':')[2]
        ensp = self.get_ensp(hgvs_source.split(':')[1].split('.')[0])
        if ensp is not None:
            hgvs_p = ensp + ':' + hgvs_source.split(':')[3]
        return hgvs_c, hgvs_p

    def get_hgvs_gel(self, variant):
        '''
        GEL variants store HGVS p dot and c dot nomenclature separately.
        This function extracts the cdot (with MANE refseq equivalent to 
        ensembl transcript ID if found) and pdot (with ensembl protein ID)
        Inputs:
            variant: (dict) dict describing single variant from JSON
        Outputs:
            hgvs_c: (str) HGVS c dot (transcript) nomenclature for the variant.
            annotated against refseq transcript ID if one exists, and ensembl
            transcript ID if there is no matched equivalent
            hgvs_p: (str) HGVS p dot (protein) nomenclature for the variant.
            this is annotated against the ensembl protein ID.
        '''
        hgvs_p = None
        hgvs_c = None
        ref_list = []
        enst_list = []
        cdnas = variant['variantAttributes']['cdnaChanges']
        protein_changes = variant['variantAttributes']['proteinChanges']

        for cdna in cdnas:
            refseq = self.convert_ensembl_to_refseq(cdna.split('(')[0])
            
            if refseq is not None:
                ref_list.append(refseq + cdna.split(')')[1])
                enst_list.append(cdna.split('(')[0])

        if len(set(ref_list)) > 1:
            raise RuntimeError(
                f"Transcript {cdnas} matched more than one MANE transcript"
            )
        elif len(set(ref_list)) == 0:
            # if there is no MANE equivalent refseq, use the ensembl
            # transcript ID, remove gene ID in brackets.
            # currently we just use the first one in the list! is there a
            # better way to do this?
            hgvs_c = re.sub("\(.*?\)", "", cdnas[0])
            ensp = self.get_ensp(cdnas[0].split('(')[0])
            for protein in protein_changes:
                if ensp in protein:
                    hgvs_p = protein
            
        else:
            hgvs_c = list(set(ref_list))[0]
            ensp = self.get_ensp(enst_list[0])
            for protein in protein_changes:
                if ensp in protein:
                    hgvs_p = protein                

        return hgvs_c, hgvs_p

    def convert_tier(self, tier, var_type):
        '''
        Convert tier to add SNV as CNVs are also Tier 1
        Inputs:
            tier (str): described GEL tiering tier for a variant in the JSON
            var_type (str): variant type
        Outputs:
            tier (str): tier for that variant converted to include variant type
            to facilitate counting a total number of variants for each tier and
            type
        '''
        if tier == "TIER1" and var_type == "SNV":
            tier = "TIER1_SNV"
        elif tier == "TIER1" and var_type == "CNV":
            tier = "TIER1_CNV"
        elif tier == "TIER2" and var_type == "SNV":
            tier = "TIER2_SNV"
        elif tier == "TIERA" and var_type == "CNV":
            tier = "TIER1_CNV"
        elif tier == "TIER1" and var_type == "STR":
            tier = "TIER1_STR"
        elif tier == "TIER2" and var_type == "STR":
            tier = "TIER2_STR"
            # TODO fix this; there are no TIER2 STRs!
            # TODO write in readme how var filtering works!
        return tier

    def get_inheritance(self, variant, proband_index):
        '''
        Work out inheritance of variant based on zygosity of parents
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            proband_index: (int) position within list of variant calls that
            belongs to the proband
        Outputs:
            inheritance (str): inferred inheritance of the variant, or None.
        '''
        # TODO entirely reconfigure this based on what GEL said!
        inheritance = None
        mother_index = None
        father_index = None
        inheritance_types = ['alternate_homozygous', 'heterozygous']
        calls = variant['variantCalls']

        if self.mother and self.father:
            # find index of variant call list for mother and father            
            for call in variant['variantCalls']:
                print(call)
                if call['participantId'] == self.mother:
                    mother_index = variant['variantCalls'].index(call)
                elif call['participantId'] == self.father:
                    father_index = variant['variantCalls'].index(call)
            if mother_index and father_index:
                if calls[proband_index]['zygosity'] in inheritance_types:

                    if (calls[
                            mother_index
                            ]['zygosity'] == 'reference_homozygous'
                        and calls[
                            father_index
                            ]['zygosity'] in inheritance_types
                        ):
                        inheritance = "paternal"

                    elif (calls[
                            mother_index
                            ]['zygosity'] in inheritance_types
                        and calls[
                            father_index
                            ]['zygosity'] == 'reference_homozygous'
                        ):
                        inheritance = "maternal"
            else:
                confusing_call = calls[proband_index]['zygosity']
                print(
                    f"proband call {confusing_call} not recognised"
                )

        # can we reliable say this? it could have been inherited from the father
        # we just don't have his data?
        elif self.mother and not self.father:
            for call in variant['variantCalls']:
                if call['participantId'] == self.mother:
                    mother_index = variant['variantCalls'].index(call)
            if calls[mother_index]['zygosity'] in inheritance_types:
                inheritance = "maternal"
        # same as above?
        elif self.father and not self.mother:
            for call in variant['variantCalls']:
                if call['participantId'] == self.father:
                    father_index = variant['variantCalls'].index(call)

            if calls[father_index]['zygosity'] in inheritance_types:
                inheritance = "paternal"

        return inheritance

    def get_str_info(self, variant, proband_index):
        '''
        Fill in variant dict for specific STR variant
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            proband_index: (int) position within list of variant calls that
            belongs to the proband
        Outputs:
            var_dict: (dict) dict of variant information extracted from JSON
            will be added to a list of dicts for conversion into dataframe.
        '''
        var_dict = self.add_columns_to_dict()
        var_dict["Chr"] = variant["coordinates"]["chromosome"]
        var_dict["Pos"] = variant["coordinates"]["start"]
        var_dict["End"] = variant["coordinates"]["end"]
        var_dict["Length"] = abs(var_dict["End"] - var_dict["Pos"])
        var_dict["Type"] = "STR"
        var_dict["Priority"] = "STR"
        var_dict["Repeat"] = variant[
            "shortTandemRepeatReferenceData"
        ]["repeatedSequence"]
        var_dict["STR1"] = variant['variantCalls'][proband_index][
            'numberOfCopies'
        ][0]['numberOfCopies']
        var_dict["STR2"] = variant['variantCalls'][proband_index][
            'numberOfCopies'
        ][1]['numberOfCopies']
        var_dict["Gene"] = self.get_gene_symbol(variant)
        return var_dict

    def get_snv_info(self, variant, pb_index, ev_index):
        '''
        Fill in variant dict for specific SNV
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            pb_index (int): index of variantCalls list for the proband
            ev_index (int): index of reportEvents list for the event
        Outputs:
            var_dict: (dict) dict of variant information extracted from JSON
            will be added to a list of dicts for conversion into dataframe.
        '''
        var_dict = self.add_columns_to_dict()
        var_dict["Chr"] = variant["variantCoordinates"]["chromosome"]
        var_dict["Pos"] = variant["variantCoordinates"]["position"]
        var_dict["Ref"] = variant["variantCoordinates"]["reference"]
        var_dict["Alt"] = variant["variantCoordinates"]["alternate"]
        var_dict["Type"] = "SNV"
        var_dict["Priority"] = self.convert_tier(
            variant["reportEvents"][ev_index]["tier"], "SNV"
        )
        var_dict["Zygosity"] = variant["variantCalls"][pb_index]["zygosity"]
        var_dict["Depth"] = variant["variantCalls"][pb_index]['depthAlternate']
        var_dict["Gene"] = self.get_gene_symbol(variant)
        var_dict['AF Max'] = self.get_af_max(variant)
        var_dict["Penetrance filter"] = variant["reportEvents"][ev_index]["penetrance"]
        var_dict["Inheritance mode"] = variant["reportEvents"][ev_index][
            "modeOfInheritance"
        ]
        var_dict["Segregation pattern"] = variant["reportEvents"][ev_index][
            "segregationPattern"
        ]
        var_dict["Inheritance"] = self.get_inheritance(variant, pb_index)
        return var_dict

    def get_cnv_info(self, variant, ev_index):
        '''
        Extract CNV info from JSON and create variant dict of required data for
        workbook
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            ev_index (int): index of reportEvents list for the event
        Outputs:
            var_dict: (dict) dict of variant information extracted from JSON
            will be added to a list of dicts for conversion into dataframe.
        '''
        var_dict = self.add_columns_to_dict()
        var_dict["Chr"] = variant["coordinates"]["chromosome"]
        var_dict["Pos"] = variant["coordinates"]["start"]
        var_dict["End"] = variant["coordinates"]["end"]
        var_dict["Length"] = abs(var_dict["End"] - var_dict["Pos"])
        var_dict["Type"] = "CNV"
        var_dict["Priority"] = self.convert_tier(
            variant["reportEvents"][ev_index]["tier"], "CNV"
        )
        var_dict["Copy Number"] = variant["variantCalls"][
            0
            ]['numberOfCopies'][0]['numberOfCopies']
        var_dict["Gene"] = self.get_gene_symbol(variant)
        var_dict['AF Max'] = self.get_af_max(variant)
        return var_dict

    def check_if_proband(self, var_calls):
        '''
        Take list of variantCalls for a variant and return index of the
        proband.
        Inputs:
            var_calls (list): list of calls of variant
        Outputs:
            index (int): index of variantCalls list for the proband
        '''
        index = None
        for call in var_calls:
            if call['participantId'] == self.proband:
                index = var_calls.index(call)
                break

        if index is None:
            raise RuntimeError(
                f"Unable to find proband ID {self.proband} in variant"
                f"{var_calls}"
            )

        return index

    def resize_variant_columns(self, sheet):
        '''
        Resize columns for both gel tiering variant page + exomiser page to the
        same width
        Inputs:
            sheet: openpxyl sheet on which to resize the columns
        '''
        for col in ['C', 'D']:
            sheet.column_dimensions[col].width = 14
        sheet.column_dimensions['M'].width = 14
        for col in ['J', 'S']:
            sheet.column_dimensions[col].width = 10
        for col in ['H', 'I', 'L']: 
            sheet.column_dimensions[col].width = 6
        for col in ['O', 'P', 'V']:
            sheet.column_dimensions[col].width = 12
        sheet.column_dimensions['B'].width = 5
        for col in ['M', 'N']:
            sheet.column_dimensions[col].width = 20
        for col in ['W', 'X']:
            sheet.column_dimensions[col].width = 25

    def create_additional_analysis_page(self):
        '''
        Get Tier3/Null SNVs for Exomiser/deNovo analysis
        '''
        variant_list = []
        ranked = []
        to_report = []
        # Look through Exomiser SNVs and return those that are ranked
        # 1, 2, or 3 and have a score >= 0.75
        for snv in self.wgs_data["interpretedGenomes"][
                self.ex_index
            ]["interpretedGenomeData"]["variants"]:
            ev_to_look_at = []
            for event in snv["reportEvents"]:
                # Filter out mitochondrial + untiered as these are likely
                # artifacts
                if (snv['variantCoordinates']['chromosome'] == 'MT' and
                    event['tier'] is None
                    ):
                    continue
                else:
                    ev_to_look_at.append(event)

            # if we have a list of non MT/untiered events, get highest
            # ranked event from this list + set it as the only report event
            # for that SNV (we do not need the other events now)
            if ev_to_look_at:
                top_event = min(ev_to_look_at, key=lambda x: float(
                    x['vendorSpecificScores']['rank']
                ))
                snv['reportEvents'] = top_event
                ranked.append(snv)

        # go through list of all variants and select top ranked and add to a
        # list of variants to report
        while len(to_report) < 3:
            rank1 = min(ranked, key=lambda x: float(
                x['reportEvents']['vendorSpecificScores']['rank']
            ))
            rank1_idx = ranked.index(rank1)
            to_report.append(rank1)
            del ranked[rank1_idx]
            rank2 = min(ranked, key=lambda x: float(
                x['reportEvents']['vendorSpecificScores']['rank']
            ))
            rank2_idx = ranked.index(rank2)
            to_report.append(rank2)
            del ranked[rank2_idx]
            rank3 = min(ranked, key=lambda x: float(
                x['reportEvents']['vendorSpecificScores']['rank']
            ))
            rank3_idx = ranked.index(rank3)
            del ranked[rank3_idx]
            to_report.append(rank3)
                    
        for snv in to_report:
            # put reportevents dict within a list to allow it to have an index
            snv['reportEvents'] = [snv['reportEvents']]
            # event index will always be 0 as we have made it so there is only
            # the top ranked event
            event_index = 0
            i = self.check_if_proband(snv["variantCalls"])
            rank = snv['reportEvents'][0]['vendorSpecificScores']['rank']
            var_dict = self.get_snv_info(snv, i, event_index)
            var_dict["Priority"] = f"Exomiser Rank {rank}"
            var_dict["HGVSc"], var_dict["HGVSp"] = self.get_hgvs_exomiser(snv)
            variant_list.append(var_dict)

        # Look through GEL variants and return those with high de novo quality
        # score
        # For SNVs
        for snv in self.wgs_data["interpretedGenomes"][
                self.gel_index
            ]["interpretedGenomeData"]["variants"]:
            for event in snv["reportEvents"]:
                if event['deNovoQualityScore'] is not None:
                    # Threshold for SNVs is 0.0013
                    if event['deNovoQualityScore'] > 0.0013:
                        event_index = snv["reportEvents"].index(event)
                        var_dict = self.get_snv_info(snv, i, event_index)
                        var_dict["Priority"] = "De novo"
                        var_dict["Inheritance"] = "De novo"
                        c_dot, p_dot = self.get_hgvs_gel(snv)
                        var_dict["HGVSc"] = c_dot
                        var_dict["HGVSp"] = p_dot
                        variant_list.append(var_dict)

        # For CNVs
        for sv in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"]["structuralVariants"]:

            for event in sv["reportEvents"]:
                # Threshold for CNVs is 0.02
                if event['deNovoQualityScore'] is not None:
                    if event['deNovoQualityScore'] > 0.02:
                        event_index = snv["reportEvents"].index(event)
                        var_dict = self.get_cnv_info(sv, event_index)
                        var_dict["Priority"] = "De novo"
                        var_dict["Inheritance"] = "De novo"
                        variant_list.append(var_dict)

        # TODO
        # handle STR / SVs which appear to be null in Exomiser (always? sometimes?)

        ex_df = pd.DataFrame(variant_list)
        ex_df = ex_df.drop_duplicates()
        if not ex_df.empty:
            # Get list of all columns except 'Priority' column
            col_except_priority = self.column_list
            col_except_priority.remove('Priority')
            # Merge exomiser df with tiered variants df, indicating if there
            # is a difference between Priority
            result = self.var_df.dtypes
            print(result)
            alt = ex_df.dtypes
            print(alt)
            ex_df = ex_df.merge(
                self.var_df,
                left_on=col_except_priority,
                right_on=col_except_priority,
                indicator='_merge',
                how='outer'
            )
            # Keep left only == keep only those that are in exomiser df and 
            # not in tiered df
            ex_df = ex_df[ex_df['_merge'] == 'left_only']
            # Clean up merge columns (drop _merge, and priority_y, rename 
            # priority_x)
            ex_df = ex_df.drop(labels=['_merge'], axis='columns')
            ex_df = ex_df.rename(columns={'Priority_x': 'Priority'})
            ex_df = ex_df.drop(labels=['Priority_y'], axis='columns')
            # Sort df by priority first, and then gene name
            ex_df = ex_df.sort_values(['Priority', 'Gene'])

        ex_df.to_excel(
            self.writer,
            sheet_name="Extended_analysis",
            index=False
        )

        self.resize_variant_columns(self.workbook["Extended_analysis"])

        # Add exomiser/de novo variant counts to summary sheet
        summary_sheet = self.workbook["Summary"]
        count_dict = {
            'B26': "Exomiser",
            'B27': "De novo",
        }
        for key, val in count_dict.items():
            if 'Priority' in ex_df.columns and val in ex_df.Priority.values:
                summary_sheet[key] = ex_df['Priority'].value_counts()[val]
            else:
                summary_sheet[key] = 0

        summary_sheet['B26'] = ex_df['Priority'].str.startswith('Exomiser').sum()

    def write_cnv_reporting_template(self, cnv_sheet_num):
        '''
        Write CNV reporting template to sheet(s) in the workbook.
        Inputs:
            cnv_sheet_num (int): number to append to CNV title
        '''
        cnv = self.workbook.create_sheet(f"cnv_interpret_{cnv_sheet_num}")
        titles ={
            "Intragenic CNVs should be analysed using the SNV guideline": [1,2],
            "Chromosomal region/gene": [3, 2],
            "Start": [3, 3],
            "Stop": [3, 4],
            "Gain/Loss": [3, 5],
            "Zygosity": [3, 6],
            "Gene Content:": [7, 2],
            "Evidence": [6, 3],
            "Possible evidence": [6, 7],
            "Prevalence in controls": [11, 2],
            "Microdel/dup syndromes": [13, 2],
            "Literature search": [15, 2],
            "Inheritance in this case": [18, 2],
            "ACMG/ACGS Classification": [20, 2],
        }

        content ={
            "Does the CNV contain protein coding genes? How many?": [8,2],
            "OMIM/green genes?": [9,2],
            "Any disease genes relevant to phenotype?": [10,2],
            "Are similar CNVs in the gnomAD-SV database? Or in DGV?": [12,2],
            "Does this CNV overlap with a known microdeletion or "\
            "microduplication syndrome? Check decipher, pubmed, new" \
            "ACMG CNV guidelines Table S3": [14,2],
            "Similar CNVs in HGMD, decipher, pubmed listed as pathogenic?"\
            "Are they de novo? Do they segregate with disease in the reported"\
            "family?": [16, 2],
            "Does gene of interest have evidence of HI/TS?": [17,2],
            "In this case is the CNV de novo, inherited, unknown? Good"\
            "phenotype fit? Non-segregation in affected family members?": [19, 2],
            "1A/3": [7, 7],
            "2G/2F for loss,\n2C-2G for gain": [11, 7],
            "2A-2B": [13, 7],
            "4A-4K": [15, 7],
            "2A, 2B, 2H-2K for loss, \n2A-2E for gain": [17, 7],
            "5A-5H": [18, 7],
            "CNV pathogenicity calculators:": [2, 9],
            "Loss:": [3, 9],
            "Gain:": [4, 9]
        }
        for key, val in titles.items():
            cnv.cell(val[0], val[1]).value = key
            cnv.cell(val[0], val[1]).font = Font(
                bold=True, name=DEFAULT_FONT.name
            )
        
        for key, val in content.items():
            cnv.cell(val[0], val[1]).value = key

        # Add pathogenicity calculator links
        cnv['J3'].hyperlink = "https://cnvcalc.clinicalgenome.org/cnvcalc/"\
            "cnv-loss"
        cnv['J4'].hyperlink = "https://cnvcalc.clinicalgenome.org/cnvcalc/"\
            "cnv-gain"

        cnv.column_dimensions['B'].width = 35
        cnv.column_dimensions['G'].width = 20
        for col in ['C', 'D', 'E', 'F']:
            cnv.column_dimensions[col].width = 15

        for i in [8, 9, 10, 12, 14, 16, 17, 19]:
            cnv.row_dimensions[i].height = 60
            cnv[f"B{i}"].alignment = Alignment(
                        wrapText=True, vertical="center"
                )
        cnv.row_dimensions[9].height = 15

        # merge evidence cells
        merge_dict = {7: 10, 11: 12, 13: 14, 15: 16, 18: 19}
        for start, end in merge_dict.items():
            cnv.merge_cells(range_string=f'G{start}:G{end}')

        for row in range(6, 20):
            cnv.merge_cells(
                start_row=row, end_row=row, start_column=3, end_column=6)

        # define which rows should have borders
        row_ranges = {
            'horizontal': [
                'B3:D3', 'B4:F4', 'B6:G6', 'B7:G7', 'B8:G8', 'B9:G9', 'B10:G10',
                'B11:G11',
                'B12:G12', 'B13:G13', 'B14:G14', 'B15:G15', 'B16:G16',
                'B17:G17', 'B18:G18', 'B19:G19', 'B20:G20',
            ],
            'horizontal_thick': [
                'B3:F3', 'B5:F5', 'B6:G6', 'B7:G7', 'B20:G20', 'B21:G21'
            ],
            'vertical': [
                'E2:E3', 'G8:G26'
            ],
            'vertical_thick': [
                'B3:B4', 'B6:B20', 'G3:G4', 'C6:C20', 'H6:H20'
            ]
        }

        self.borders(row_ranges, cnv)

        # add some colour
        colour_cells = {
            'FFC000': ['G17'],
            'FFFF00': ['G15', 'G18'],
            '0070C0': ['G13'],
            '00B050': ['G11'],
            'D9D9D9': ['G7']
        }
        self.colours(colour_cells, cnv)

        # align text
        for row in [7, 11, 13, 15, 17, 18]:
            cnv[f"G{row}"].alignment = Alignment(
                wrapText=True, vertical="center", horizontal="center"
            )

    def write_reporting_template(self, report_sheet_num) -> None:
        """
        Writes sheet(s) to Excel file with formatting for reporting against
        ACMG criteria

        """
        report = self.workbook.create_sheet(
            f"snv_interpret_{report_sheet_num}"
        )

        titles = {
            "Gene": [2, 2],
            "HGVSc": [2, 3],
            "HGVSp": [2, 4],
            "EVIDENCE": [8, 3],
            "PATHOGENIC": [8, 7],
            "P_STRENGTH": [8, 8],
            "P_POINTS": [8, 9],
            "BENIGN": [8, 10],
            "B_STRENGTH": [8, 11],
            "B_POINTS": [8, 12],
            "Associated disease": [4, 2],
            "Known inheritance": [5, 2],
            "Prevalence": [6, 2],
            ("Allele frequency is >5% (or gene-specific cut off) in "
             "population data e.g. gnomAD, UKB"): [9, 2],
            ("Null variant in a gene where LOF is known mechanism "
             "of disease\nand non-canonical splice variants where "
             "RNA analysis confirms\naberrant transcription"): [10, 2],
            ("Same AA change as previously established pathogenic "
             "variant\nregardless of nucleotide change and splicing "
             "variants within\nsame motif with identical predicted "
             "effect"): [11, 2],
            ("De novo (confirmed) / observed in\nhealthy adult "
             "with full penetrance expected at an early age"): [12, 2],
            "In vivo / in vitro functional studies": [13, 2],
            "Prevalence in affected > controls": [14, 2],
            ("In mutational hot spot and/or critical functional "
             "domain, without\nbenign variation"): [15, 2],
            ("Freq in controls eg gnomAD, low/absent (PM2) or allele "
             "frequency is greater than expected for disorder (BS1)"): [16, 2],
            "Detected in trans/in cis with pathogenic variant": [17, 2],
            ("In frame protein length change/stop-loss variants, "
             "non repeat\nvs. repeat region"): [18, 2],
            ("Missense change at AA where different likely/pathogenic\n"
             "missense change seen before"): [19, 2],
            "Assumed de novo (no confirmation)": [20, 2],
            "Cosegregation with disease in family, not in unaffected": [21, 2],
            ("Missense where low rate of benign missense and common\n"
             "mechanism (Z score ≥3.09), or missense where LOF common\n"
             "mechanism"): [22, 2],
            "Multiple lines of computational evidence": [23, 2],
            ("Phenotype/FH specific for disease of single etiology, or\n"
             "alternative genetic cause of disease detected"): [24, 2],
            ("Synonymous change, no affect on splicing, not conserved; "
             "splice\nvariants confirmed to have no impact"): [25, 2],
            "POINTS": [26, 7],
        }
        for key, val in titles.items():
            report.cell(val[0], val[1]).value = key
            report.cell(val[0], val[1]).font = Font(
                bold=True, name=DEFAULT_FONT.name
            )
        classifications = {
            "PVS1": [(10, 7)],
            "PS1": [(11, 7)],
            "PS2": [(12, 7)],
            "PS3": [(13, 7)],
            "PS4": [(14, 7)],
            "PM1": [(15, 7)],
            "PM2": [(16, 7)],
            "PM3": [(17, 7)],
            "PM4": [(18, 7)],
            "PM5": [(19, 7)],
            "PM6": [(20, 7)],
            "PP1": [(21, 7)],
            "PP2": [(22, 7)],
            "PP3": [(23, 7)],
            "PP4": [(24, 7)],
            "BA1": [(9, 10)],
            "BS2": [(12, 10)],
            "BS3": [(13, 10)],
            "BS1": [(16, 10)],
            "BP2": [(17, 10)],
            "BP3": [(18, 10)],
            "BS4": [(21, 10)],
            "BP1": [(22, 10)],
            "BP4": [(23, 10)],
            "BP5": [(24, 10)],
            "BP7": [(25, 10)]
        }

        for key, values in classifications.items():
            for val in values:
                report.cell(val[0], val[1]).value = key

        # nice formatting of title text and columns
        for col in (['B', 'C', 'G', 'H', 'I', 'J', 'K', 'L']):
            for row in range(8, 26):
                if row == 8:
                    report[f"{col}{row}"].alignment = Alignment(
                           wrapText=True, vertical="center",
                           horizontal="center"
                    )
                elif (col == "B" and row != 8) or (col == "C" and row != 8):
                    report.row_dimensions[row].height = 50
                    report[f"{col}{row}"].alignment = Alignment(
                           wrapText=True, vertical="center"
                    )
                else:
                    report[f"{col}{row}"].alignment = Alignment(
                           wrapText=True, vertical="center",
                           horizontal="center"
                    )
                    report[f"{col}{row}"].font = Font(size=14,
                                                      name=DEFAULT_FONT.name)

        for col in (['B', 'C', 'D']):
            for row in range(2, 7):
                report.row_dimensions[row].height = 20
                report[f"{col}{row}"].font = Font(size=14, bold=True,
                                                  name=DEFAULT_FONT.name)

        # merge associated disease, inheritance and prevalence cells
        for row in range(4, 8):
            report.merge_cells(
                start_row=row, end_row=row, start_column=3, end_column=12)

        # merge evidence cells
        for row in range(8, 27):
            report.merge_cells(
                start_row=row, end_row=row, start_column=3, end_column=6)

        # merge POINTS cells
        report.merge_cells(
                start_row=26, end_row=26, start_column=8, end_column=12)

        # set appropriate widths
        report.column_dimensions['B'].width = 62
        report.column_dimensions['C'].width = 35
        report.column_dimensions['D'].width = 35
        report.column_dimensions['E'].width = 5
        report.column_dimensions['F'].width = 5
        report.column_dimensions['G'].width = 14
        report.column_dimensions['H'].width = 14
        report.column_dimensions['I'].width = 14
        report.column_dimensions['J'].width = 14
        report.column_dimensions['K'].width = 14
        report.column_dimensions['L'].width = 14

        # do some colouring
        colour_cells = {
            'E46C0A': ['G11', 'G12', 'G13', 'G14'],
            'FFC000': ['G15', 'G16', 'G17', 'G18', 'G19', 'G20'],
            'FFFF00': ['G21', 'G22', 'G23', 'G24'],
            '00B0F0': ['J12', 'J13', 'J16', 'J21'],
            '92D050': ['J17', 'J18', 'J22', 'J23', 'J24', 'J25'],
            '0070C0': ['J9'],
            'FF0000': ['G10'],
            'D9D9D9': ['G9', 'G25', 'H9', 'H25', 'I9', 'I25',
                       'J10', 'J11', 'J14', 'J15', 'J19', 'J20',
                       'K10', 'K11', 'K14', 'K15', 'K19', 'K20',
                       'L10', 'L11', 'L14', 'L15', 'L19', 'L20'
                       ]

        }

        self.colours(colour_cells, report)

        # add some borders
        row_ranges = {
            'horizontal': [
                'B3:D3', 'B4:J4', 'B5:L5',
                'B6:L6', 'B7:L7', 'B8:L8', 'B9:L9', 'B10:L10', 'B11:L11',
                'B12:L12', 'B13:L13', 'B14:L14', 'B15:L15', 'B16:L16',
                'B17:L17', 'B18:L18', 'B19:L19', 'B20:L20', 'B21:L21',
                'B22:L22', 'B23:L23', 'B24:L24', 'B25:L25'
            ],
            'horizontal_thick': [
                'B2:D2', 'B4:L4', 'B7:L7', 'B8:L8', 'B26:L26', 'B27:L27'
            ],
            'vertical': [
                'E2:E3', 'G8:G26', 'H8:H25', 'I8:I25', 'J8:J25',
                'K8:K25', 'L8:L25',
            ],
            'vertical_thick': [
                'B2:B6', 'B8:B26', 'C2:C6', 'C8:C26', 'M4:M6', 'M8:M26',
                'E2:E3'
            ]
        }
        self.borders(row_ranges, report)

    def colours(self, colour_cells, sheet):
        '''
        Add colour to cells in workbook.
        Inputs:
            colour_cells (dict): dict of colour number keys with cells that
            should be that colour as values
            sheet (str): sheet in workbook to add colour to
        '''
        for colour, cells in colour_cells.items():
            for cell in cells:
                sheet[cell].fill = PatternFill(
                    patternType="solid", start_color=colour
                )

    def borders(self, row_ranges, sheet):
        '''
        Add borders to sheet.
        Inputs:
            row_ranges (dict): dict of border type and rows/columns that should
            have that border
            sheet (str): sheet in workbook to add borders to
        '''
        for side, values in row_ranges.items():
            for row in values:
                for cells in sheet[row]:
                    for cell in cells:
                        # border style is immutable => copy current and modify
                        cell_border = cell.border.copy()
                        if side == 'horizontal':
                            cell_border.top = THIN
                        if side == 'horizontal_thick':
                            cell_border.top = MEDIUM
                        if side == 'vertical':
                            cell_border.left = THIN
                        if side == 'vertical_thick':
                            cell_border.left = MEDIUM
                        cell.border = cell_border
        # if self.args.lock_sheet:
        #     cell_to_unlock = ["B3", "C3", "D3", "C4", "C5", "C6",
        #                       "C9", "C10", "C11", "C12", "C13", "C14", "C15",
        #                       "C16", "C17", "C18", "C19", "C20", "C21", "C22",
        #                       "C23", "C24", "C25", "C26", "H10", "H11",
        #                       "H12", "H13", "H14", "H15", "H16", "H17", "H18",
        #                       "H19", "H20", "H21", "H22", "H23", "H24", "I10",
        #                       "I11", "I12", "I13", "I14", "I15", "I16", "I17",
        #                       "I18", "I19", "I20", "I21", "I22", "I23", "I24",
        #                       "K9", "K12", "K13", "K16", "K17", "K18", "K21",
        #                       "K22", "K23", "K24", "K25", "L9", "L12", "L13",
        #                       "L16", "L17", "L18", "L21", "L22", "L23", "L24",
        #                       "L25", "H26"]
            # self.lock_sheet(ws=report,
            #                 cell_to_unlock=cell_to_unlock,
            #                 start_row=report.max_row,
            #                 start_col=report.max_column,
            #                 unlock_row_num=ROW_TO_UNLOCK,
            #                 unlock_col_num=COL_TO_UNLOCK)

    def drop_down(self) -> None:
        """
        Function to add drop-downs in the report tab for entering
        ACMG criteria for classification, as well as a boolean
        drop down into the additional 'Interpreted' column of
        the variant sheet(s).
        """
        wb = load_workbook(filename=self.args.output)

        # adding dropdowns in report table
        for sheet_num in range(1, self.args.acmg+1):
            # adding strength dropdown except for BA1
            report_sheet = wb[f"snv_interpret_{sheet_num}"]
            cells_for_strength = ['H10', 'H11', 'H12', 'H13', 'H14', 'H15',
                                  'H16', 'H17', 'H18', 'H19', 'H20', 'H21',
                                  'H22', 'H23', 'H24', 'K12', 'K13', 'K16',
                                  'K17', 'K18', 'K21', 'K22', 'K23', 'K24',
                                  'K25']
            strength_options = '"Very Strong, Strong, Moderate, \
                                 Supporting, NA"'
            self.get_drop_down(dropdown_options=strength_options,
                               prompt='Select from the list',
                               title='Strength',
                               sheet=report_sheet,
                               cells=cells_for_strength)

            # add stregth for BA1
            BA1_options = '"Stand-Alone, Very Strong, Strong, Moderate, \
                            Supporting, NA"'
            self.get_drop_down(dropdown_options=BA1_options,
                               prompt='Select from the list',
                               title='Strength',
                               sheet=report_sheet,
                               cells=['K9'])

            # adding final classification dropdown
            report_sheet['B26'] = 'FINAL ACMG CLASSIFICATION'
            report_sheet['B26'].font = Font(bold=True, name=DEFAULT_FONT.name)
            class_options = '"Pathogenic,Likely Pathogenic, \
                              Uncertain Significance, \
                              Likely Benign, Benign"'
            self.get_drop_down(dropdown_options=class_options,
                               prompt='Select from the list',
                               title='ACMG classification',
                               sheet=report_sheet,
                               cells=['C26'])

        # adding Interpreted column dropdown in the first variant sheet tab
        # first_variant_sheet = wb[self.args.sheets[0]]
        # interpreted_options = '"YES,NO"'
        # col_letter = self.get_col_letter(first_variant_sheet, "Interpreted")
        # num_variant = self.vcfs[0].shape[0]
        # if num_variant > 0:
        #     cells_for_variant = []
        #     for i in range(num_variant):
        #         cells_for_variant.append(f"{col_letter}{i+2}")
        #     self.get_drop_down(dropdown_options=interpreted_options,
        #                        prompt='Choose YES or NO',
        #                        title='Variant interpreted or not?',
        #                        sheet=first_variant_sheet,
        #                        cells=cells_for_variant)
        wb.save(self.args.output)

    def get_drop_down(self, dropdown_options, prompt, title, sheet, cells) -> None:
        """
        create the drop-downs items for designated cells

        Parameters
        ----------
        dropdown_options: str
            str containing drop-down items
        prompt: str
            prompt message for drop-down
        title: str
            title message for drop-down
        sheet: openpyxl.Writer writer object
            current worksheet
        cells: list
            list of cells to add drop-down
        """
        options = dropdown_options
        val = DataValidation(type='list', formula1=options,
                             allow_blank=True)
        val.prompt = prompt
        val.promptTitle = title
        sheet.add_data_validation(val)
        for cell in cells:
            val.add(sheet[cell])
        val.showInputMessage = True
        val.showErrorMessage = True
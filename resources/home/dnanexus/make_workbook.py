import dxpy
import networkx
import gzip
import json
import obonet
import openpyxl.drawing
import openpyxl.drawing.image
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, DEFAULT_FONT, Font
import os
from pathlib import Path
import re

from excel_styles import ExcelStyles, DropDown
from get_variant_info import VariantInfo, VariantNomenclature

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
        self.proband = None
        self.proband_sex = None
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
            "Zygosity",
            "Inheritance mode",
            "Inheritance",
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
        self.index_interpretation_services()
        self.create_gel_tiering_variant_page()
        self.create_additional_analysis_page()
        self.str_image_page()
        self.writer.close()
        if self.args.acmg:
            for i in range(1, self.args.acmg+1):
                self.write_snv_reporting_template(i)
        if self.args.cnv:
            for i in range(1, self.args.cnv+1):
                self.write_cnv_reporting_template(i)
        self.workbook.save(self.args.output)
        if self.args.acmg:
            DropDown.drop_down(self)
        print('Done!')

    def open_files(self):
        '''
        Open input files and read into variables.
        '''
        with open(self.args.json) as f:
            self.wgs_data = json.load(f)

        with gzip.open(self.args.mane_file) as f:
            self.mane = [x.decode('utf8').strip() for x in f.readlines()]

        with open(self.args.refseq_tsv) as refseq_tsv:
            self.refseq_tsv = refseq_tsv.readlines()

    def summary_page(self):
        '''
        Add summary page. Create a page in the workbook to populate with
        details about the case and variants for interpretation.
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
        '''
        summary_sheet = self.workbook.create_sheet("Summary")

        self.bold_content = {
            (1, 1): "Family ID",
            (6, 1): "Proband",
            (7, 1): "Mother",
            (8, 1): "Father",
            (5, 2): "ID",
            (5, 3): "GM/SP number",
            (5, 4): "NUH number",
            (5, 5): "Sex",
            (5, 6): "Affected?",
            (5, 7): "HPO",
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
            (21, 2): "In this case",
            (21, 3): "To be reported",
            (22, 1): "SNV Tier 1",
            (23, 1): "SNV Tier 2",
            (24, 1): "CNV Tier 1",
            (25, 1): "STR Tier 1",
            (27, 1): "Exomiser top 3 (score ≥ 0.75)",
            (28, 1): "De novo",
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
        # TODO: Method to get SP number/NUH number from Epic via Clarity
        # export and autopopulate summary page with this info.

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
        for col in ["C", "D", "F"]:
            summary_sheet.column_dimensions[col].width = 14

        row_ranges = {
            'horizontal': [
                'A33:C33', 'A34:C34', 'A36:C36', 'A30:B30', 'A31:B31',
                'A32:B32'
            ],
            'vertical': [
                'A33:A35', 'B33:B35', 'C33:C35', 'D33:D35', 'A30:A31',
                'B30:B31', 'C30:C31',
            ]
        }

        ExcelStyles.borders(self, row_ranges, summary_sheet)

    def get_hpo_obo(self):
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
        version = self.wgs_data['referral']['referral_data']["pedigree"][
            "members"
            ][0]["hpoTermList"][0]['hpoBuildNumber']

        if self.args.obo_files:
            if version == "v2019_02_12":
                obo = "/home/dnanexus/obo_files/hpo_v20190212.obo"
            elif version == "releases/2018-10-09":
                obo = "/home/dnanexus/obo_files/hpo_v20181009.obo"
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

    def get_hpo_terms(self, member, obo):
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

    def add_person_data_to_summary(self, member, index, obo):
        '''
        Function to add data that is added for all family members to the
        summary sheet. Isolated here as it is the same for all family members
        Inputs:
            member (dict): dictionary for that member in the pedigree in the
            GEL JSON
            obo (str): path to obo file
            index (int): row to populate in the sheet, specific to that family
            member
        Outputs:
            None, adds content to openpxyl workbook
        '''
        self.summary_content[(index, 2)] = member["participantId"]
        self.summary_content[(index, 5)] = member["sex"]
        self.summary_content[(index, 6)] = member["affectionStatus"]
        self.summary_content[(index, 7)] = self.get_hpo_terms(member, obo)

    def person_data(self):
        '''
        Find data for participants and add to summary sheet.
        This function will find the proband and add their affected status, HPO
        term names, participant ID, sample ID to the summary sheet
        It will also do this for any relatives in the JSON
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
        '''
        # Get version of HPO to use for terms
        obo = self.get_hpo_obo()
        pb_relate = lambda x: x["additionalInformation"]["relation_to_proband"]

        for member in self.wgs_data['referral']['referral_data']["pedigree"][
            "members"
            ]:
            if member["isProband"]:
                self.add_person_data_to_summary(member, 6, obo)
                self.proband = member["participantId"]
                self.proband_sex = member["sex"]
                self.summary_content[(10, 2)] = member["samples"][0][
                    "sampleId"
                ]

            elif pb_relate(member) == "Mother":
                self.add_person_data_to_summary(member, 7, obo)
                self.mother = member["participantId"]

            elif pb_relate(member) == "Father":
                self.add_person_data_to_summary(member, 8, obo)
                self.father = member["participantId"]

            else:
                self.summary_content[(9, 1)] = pb_relate(member)
                self.add_person_data_to_summary(member, 9, obo)

    def get_panels(self):
        '''
        Function to get panels from JSON and add to summary content dict to
        then add to summary content sheet
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
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
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
        '''
        for penetrance in self.wgs_data['referral']['referral_data'][
            'pedigree'
            ]['diseasePenetrances']:
            if penetrance['specificDisease'] == self.summary_content[(2,2)]:
                disease_penetrance = penetrance["penetrance"]

        self.summary_content[(3, 2)] = disease_penetrance

    def index_interpretation_services(self):
        '''
        The JSON contains two interpretation services: GEL tiering and Exomiser
        This function finds the index for each so these can be referred to
        correctly and sets self.gel_index to the index for the GEL tiering and
        self.ex_index to to the index for the Exomiser interpretation.
        Inputs:
            None
        Outputs:
            None, sets indexs for GEL tiering and exomiser in the JSON.
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

    def str_image_page(self):
        '''
        STR table is useful for interpretation, so will be included on a sheet
        so it can be referred to during interpretation.
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
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

    def create_gel_tiering_variant_page(self):
        '''
        Take variants from GEL tiering JSON and format into sheet in Excel
        workbook.
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
        '''
        variant_list = []

        # SNVs
        for snv in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"]["variants"]:
            for event in snv["reportEvents"]:
                if event["tier"] in ["TIER1", "TIER2"]:
                    event_index = snv["reportEvents"].index(event)
                    var_dict = VariantInfo.get_snv_info(
                        snv,
                        self.proband,
                        event_index,
                        self.column_list,
                        self.mother,
                        self.father,
                        self.proband_sex
                        )
                    c_dot, p_dot = VariantNomenclature.get_hgvs_gel(
                        snv,
                        self.mane,
                        self.refseq_tsv
                        )
                    var_dict["HGVSc"] = c_dot
                    var_dict["HGVSp"] = p_dot
                    variant_list.append(var_dict)

        # STRs
        for str in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"][
                "shortTandemRepeats"
            ]:
            for event in str["reportEvents"]:
                if event["tier"] == "TIER1":
                    var_dict = VariantInfo.get_str_info(
                        str, self.proband, self.column_list
                    )
                    variant_list.append(var_dict)

        # CNVs
        for cnv in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"]["structuralVariants"]:
            for event in cnv["reportEvents"]:
                event_index = cnv["reportEvents"].index(event)
                # CNVs can be reported as Tier 1 or Tier A, GEL updated the
                # nomenclature in 2024
                if cnv["reportEvents"][event_index]["tier"] in [
                    "TIER1", "TIERA"
                    ]:
                    var_dict = VariantInfo.get_cnv_info(
                        cnv, event_index, self.column_list
                    )
                    variant_list.append(var_dict)

        # Add all variants into dataframe
        self.var_df = pd.DataFrame(variant_list)
        self.var_df = self.var_df.drop_duplicates()
        self.var_df['Depth'] = self.var_df['Depth'].astype(object)

        # if df is not empty
        if not self.var_df.empty:
            # if both Tier 1 and Tier 2, keep Tier 1 entry only
            self.var_df = self.var_df.sort_values(
                by=['Priority']
                ).drop_duplicates(subset=['Chr', 'Pos', 'Ref', 'Alt', 'End'])
            # Sort by Priority and then Gene symbol
            self.var_df = self.var_df.sort_values(['Priority', 'Gene'])
            self.var_df.to_excel(
                self.writer, sheet_name="Variants", index=False
            )

        # Set column widths
        ExcelStyles.resize_variant_columns(self, self.workbook["Variants"])

        # Add counts to summary sheet
        summary_sheet = self.workbook["Summary"]
        count_dict = {
            'B22': "TIER1_SNV",
            'B23': "TIER2_SNV",
            'B24': "TIER1_CNV",
            'B25': "TIER1_STR",
        }
        for key, val in count_dict.items():
            if val in self.var_df.Priority.values:
                summary_sheet[key] = self.var_df[
                    'Priority'
                ].value_counts()[val]
            else:
                summary_sheet[key] = 0

    def create_additional_analysis_page(self):
        '''
        Get Tier3/Null SNVs for Exomiser/deNovo analysis
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
        '''
        variant_list = []
        ranked = []
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
                top_event = min(ev_to_look_at, key=lambda x:
                    x['vendorSpecificScores']['rank']
                )
                snv['reportEvents'] = top_event
                ranked.append(snv)

        to_report = VariantInfo.get_top_3_ranked(ranked)

        for snv in to_report:
            # put reportevents dict within a list to allow it to have an index
            snv['reportEvents'] = [snv['reportEvents']]
            # event index will always be 0 as we have made it so there is only
            # the top ranked event
            event_index = 0
            rank = snv['reportEvents'][0]['vendorSpecificScores']['rank']
            var_dict = VariantInfo.get_snv_info(
                snv,
                self.proband,
                event_index,
                self.column_list,
                self.mother,
                self.father,
                self.proband_sex
            )
            var_dict["Priority"] = f"Exomiser Rank {rank}"
            var_dict["HGVSc"], var_dict["HGVSp"] = (
                VariantNomenclature.get_hgvs_exomiser(
                    snv,
                    self.mane,
                    self.refseq_tsv)
                )
            variant_list.append(var_dict)

        # Look through GEL variants and return those with high de novo quality
        # score
        # TODO could these have already been reported in the gel tiering sheet?
        # if so need to avoid duplicating!
        # For SNVs
        for snv in self.wgs_data["interpretedGenomes"][
                self.gel_index
            ]["interpretedGenomeData"]["variants"]:
            for event in snv["reportEvents"]:
                if event['deNovoQualityScore'] is not None:
                    # Threshold for SNVs is 0.0013
                    if event['deNovoQualityScore'] > 0.0013:
                        event_index = snv["reportEvents"].index(event)
                        var_dict = VariantInfo.get_snv_info(
                            snv,
                            self.proband,
                            event_index,
                            self.column_list,
                            self.mother,
                            self.father,
                            self.proband_sex
                        )
                        var_dict["Priority"] = "De novo"
                        var_dict["Inheritance"] = "De novo"
                        var_dict["HGVSc"], var_dict["HGVSp"] = (
                            VariantNomenclature.get_hgvs_gel(
                                snv,
                                self.mane,
                                self.refseq_tsv)
                        )
                        variant_list.append(var_dict)

        # CNVs
        for sv in self.wgs_data["interpretedGenomes"][
            self.gel_index
            ]["interpretedGenomeData"]["structuralVariants"]:
            for event in sv["reportEvents"]:
                # Threshold for CNVs is 0.02
                if event['deNovoQualityScore'] is not None:
                    if event['deNovoQualityScore'] > 0.02:
                        event_index = snv["reportEvents"].index(event)
                        var_dict = VariantInfo.get_cnv_info(sv,
                                                            event_index,
                                                            self.column_list)
                        var_dict["Priority"] = "De novo"
                        var_dict["Inheritance"] = "De novo"
                        variant_list.append(var_dict)

        # TODO
        # handle STR / SVs which appear to be null in Exomiser
        # (always? sometimes?)

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

        ExcelStyles.resize_variant_columns(
            self, self.workbook["Extended_analysis"]
        )

        # Add exomiser/de novo variant counts to summary sheet
        summary_sheet = self.workbook["Summary"]
        if 'Priority' in ex_df.columns:
            summary_sheet['B28'] = ex_df['Priority'].str.startswith(
                "De novo"
            ).sum()
            summary_sheet['B27'] = ex_df['Priority'].str.startswith(
                'Exomiser'
            ).sum()

    def write_cnv_reporting_template(self, cnv_sheet_num):
        '''
        Write CNV reporting template to sheet(s) in the workbook.
        Inputs:
            cnv_sheet_num (int): number to append to CNV title
        Outputs:
            None, adds content to openpxyl workbook
        '''
        cnv = self.workbook.create_sheet(f"cnv_interpret_{cnv_sheet_num}")
        titles = {
            "Intragenic CNVs should be analysed using SNV guidelines": [1,2], 
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

        content = {
            "Does the CNV contain protein coding genes? How many?": [8,2], 
            "OMIM/green genes?": [9,2], 
            "Any disease genes relevant to phenotype?": [10,2], 
            "Are similar CNVs in the gnomAD-SV database? Or in DGV?": [12,2], 
            "Does this CNV overlap with a known microdeletion or "
            "microduplication syndrome? Check decipher, pubmed, new " 
            "ACMG CNV guidelines Table S3": [14,2], 
            "Similar CNVs in HGMD, decipher, pubmed listed as pathogenic?"
            "Are they de novo? Do they segregate with disease in the reported"
            "family?": [16, 2], 
            "Does gene of interest have evidence of HI/TS?": [17,2], 
            "In this case is the CNV de novo, inherited, unknown? Good "
            "phenotype fit? Non-segregation in affected family"
            "members?": [19, 2],
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
                'B3:D3', 'B4:F4', 'B6:G6', 'B7:G7', 'B8:G8', 'B9:G9',
                'B10:G10', 'B11:G11',
                'B12:G12', 'B13:G13', 'B14:G14', 'B15:G15', 'B16:G16',
                'B17:G17', 'B18:G18', 'B19:G19', 'B20:G20',
            ],
            'horizontal_thick': [
                'B3:F3', 'B5:F5', 'B6:G6', 'B7:G7', 'B20:G20', 'B21:G21'
            ],
            'vertical': [
                'E2:E3', 'G6:G20'
            ],
            'vertical_thick': [
                'B3:B4', 'B6:B20', 'G3:G4', 'C6:C20', 'H6:H20'
            ]
        }

        ExcelStyles.borders(self, row_ranges, cnv)

        # add some colour
        colour_cells = {
            'FFC000': ['G17'],
            'FFFF00': ['G15', 'G18'],
            '0070C0': ['G13'],
            '00B050': ['G11'],
            'D9D9D9': ['G7']
        }
        ExcelStyles.colours(self, colour_cells, cnv)

        # align text
        for row in [7, 11, 13, 15, 17, 18]:
            cnv[f"G{row}"].alignment = Alignment(
                wrapText=True, vertical="center", horizontal="center"
            )

    def write_snv_reporting_template(self, report_sheet_num) -> None:
        """
        Writes sheet(s) to Excel file with formatting for reporting against
        ACMG criteria
        Inputs:
            None
        Outputs:
            None, adds content to openpxyl workbook
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

        ExcelStyles.colours(self, colour_cells, report)

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
        ExcelStyles.borders(self, row_ranges, report)

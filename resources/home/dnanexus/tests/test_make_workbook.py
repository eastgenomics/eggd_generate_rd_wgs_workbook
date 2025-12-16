#!/usr/bin/env python3
import argparse
import pandas as pd
import pytest
import os
import obonet
import sys
import json
import warnings
import excel_styles
from openpyxl import Workbook
from make_workbook import excel
import get_variant_info as var_info
from start_process import SortArgs
from unittest import mock
from unittest.mock import MagicMock, patch


class TestWorkbook():
    '''
    Tests for excel() class in make_workbook script
    '''
    summary_content = {}
    panels = {
        '486': {
            'rcode': 'R123',
            'panel_name': 'Paediatric disorders'
        }
    }
    wgs_data = {
        "family_id": ["r12345"],
        "interpretation_request_data": {
            "json_request": {
                "pedigree": {
                    "members": [
                        {
                            "hpoTermList": [
                                {"hpoBuildNumber": "vXXXXXX"}
                            ]
                        },
                        {
                            "participantId": "proband_id",
                            "sex": "MALE",
                            "isProband": True
                        }
                    ],
                    "diseasePenetrances": [
                        {
                            "penetrance": "complete",
                            "specificDisease": "Congenital malformation"
                        },
                        {
                            "penetrance": "incomplete",
                            "specificDisease": "OtherDisease"
                        }
                    ],
                    "analysisPanels": [
                        {
                            "panelId": "486",
                            "panelName": "286",
                            "specificDisease": "Congenital malformation",
                            "panelVersion": "2.2"
                        }
                    ]
                }
            }
        },
        "interpretedGenomes": [
            {
                "interpretedGenomeData": {
                    "interpretationService": "genomics_england_tiering",
                    "variants": [
                        {
                            "variantCoordinates": {
                                "chromosome": "1",
                                "position": 12345,
                                "reference": "A",
                                "alternate": "G"
                            },
                            "variantCalls": [
                                {
                                    "participantId": "proband_id",
                                    "zygosity": "alternate_homozygous",
                                    "depthReference": 20,
                                    "depthAlternate": 29
                                }
                            ],
                            "variantAttributes": {
                                "alleleFrequencies": [],
                                "additionalTextualVariantAnnotations": {
                                    "hgvs": ["TEST:c.123A>G"]
                                },
                                "cdnaChanges": ["TEST:c.123A>G"],
                                "proteinChanges": ["TEST:p.Arg123Gly"]
                            },
                            "reportEvents": [
                                {
                                    "tier": "TIER1",
                                    "genomicEntities": [
                                        {"geneSymbol": "TESTGENE", "type": "gene"}
                                    ],
                                    "penetrance": "incomplete",
                                    "modeOfInheritance": "None"
                                }
                            ]
                        }
                    ],
                    "shortTandemRepeats": [],
                    "structuralVariants": []
                }
            }
        ]
    }


    def test_get_panels_extracts_data_from_input_panel_json(self):
        '''
        '486' and 'Congenital malformation' are sourced directly from the GEL
        JSON and 'R123' and 'Paediatric disorders (2.2)' are sourced from
        looking the panel ID up in the panel JSON
        '''
        excel.get_panels(self)
        assert self.summary_content == {
            (14, 1): '486',
            (2, 2): 'Congenital malformation',
            (14, 2): 'R123',
            (14, 3): 'Paediatric disorders (2.2)'
        }

    def test_get_penetrance(self):
        '''
        Check that penetrance is extracted from JSON as expected, and matched
        to the specific disease(s) in the referral
        '''
        excel.get_penetrance(self)
        assert self.summary_content[(3,2)] == "complete, incomplete"

    @mock.patch('pandas.read_csv')
    def test_epic_extract_with_incorrect_column_names_raises_error(self, pd_read_csv_mock):
        self.args = argparse.Namespace
        self.args.epic_clarity = None
        self.other_relation = False
        # This should error as required Specimen Identifier cols are missing
        mock_df = pd.DataFrame(
            {
            "Year of Birth": [1937, 1975],
            "Patient Stated Gender": [1, 2],
            "WGS Referral ID": ["r12345", "r67890"]
            }
        )
        pd_read_csv_mock.return_value = mock_df
        with pytest.raises(ValueError):
            excel.add_epic_data(self)

    def test_alternative_homozygous_notation_becomes_homozygous(self):
        """
        Test that 'alternate_homozygous' notation in zygosity is converted to
        'homozygous' in the workbook
        """

        # Set up excel instance with mocked dependencies
        mock_args = MagicMock()
        excel_instance = excel(mock_args)
        excel_instance.wgs_data = self.wgs_data
        excel_instance.proband = "proband_id"
        excel_instance.proband_sex = "MALE"
        excel_instance.mane = []
        excel_instance.refseq_tsv = []
        excel_instance.var_df = pd.DataFrame()

        # Call helper fnctions that performs normalisation
        excel_instance.get_interpreted_genome_format()
        excel_instance.index_interpretation_services()

        # Patch whats needed for the excel_instance
        with patch.object(pd.DataFrame, "to_excel", return_value=None), \
            patch.object(excel_instance, "open_files"), \
            patch.object(excel_instance, "writer", create=True), \
            patch.object(excel_instance, "workbook", create=True), \
            patch("excel_styles.DropDown.drop_down"), \
            patch("excel_styles.ExcelStyles.borders"):

            excel_instance.writer = MagicMock()
            excel_instance.workbook = MagicMock()

            # Call function that changes"alternate_homozygous" to "homozygous"
            excel_instance.create_gel_tiering_variant_page()

            # Check alternate_homozygous changed to homozygous in mock workbook
            assert "homozygous" in excel_instance.var_df["Zygosity"].values
            assert "alternate_homozygous" not in excel_instance.var_df["Zygosity"].values

    def test_write_snv_report_colouring(self):
        '''
        Test that the function to add SNV reporting colouring to the workbook
        runs without error.
        '''
        wb = Workbook()
        wb.remove(wb.active)
        writer = excel(None)
        writer.workbook = wb

        writer.write_snv_reporting_template(1)
        sheet = wb["snv_interpret_1"]

        # Check G2 and G3 are yellow
        assert sheet["G2"].fill.start_color.rgb in ("FFFF00", "00FFFF00")
        assert sheet["G3"].fill.start_color.rgb in ("FFFF00", "00FFFF00")

        # Check all of column M is yellow
        assert all(
            sheet[f"M{row}"].fill.start_color.rgb in ("FFFF00", "00FFFF00")
            for row in range(1, sheet.max_row + 1)
        )

    def test_cnv_report_colouring(self):
        '''
        Test that the function to add CNV reporting colouring to the workbook
        runs without error.
        '''
        wb = Workbook()
        wb.remove(wb.active)
        writer = excel(None)
        writer.workbook = wb

        writer.write_cnv_reporting_template(1)
        sheet = wb["cnv_interpret_1"]


        # Check all of column H is yellow
        assert all(
            sheet[f"H{row}"].fill.start_color.rgb in ("FFFF00", "00FFFF00")
            for row in range(1, sheet.max_row + 1)
        )
class TestInterpretationService():
    '''
    Test that the function to find interpretation service works as expected
    '''
    genome_format = None
    genome_data_format = None
    ex_index = None
    gel_index = None

    wgs_data = {
        'interpretedGenomes': [
            {'interpretedGenomeData': {
                'interpretationService': 'genomics_england_tiering'}
            },
            {'interpretedGenomeData': {'interpretationService': 'exomiser'}}
        ]
    }

    def test_that_camelcase_format_is_found(self):
        '''
        Test that the function get_interpreted_genome_format returns the
        correct genome_format and genome_data_format for camelcase fields in
        JSON
        '''
        excel.get_interpreted_genome_format(self)

        assert (
            self.genome_format == 'interpretedGenomes' and
            self.genome_data_format == 'interpretedGenomeData'
        )

    def test_indexing_of_interpretation_service(self):
        '''
        Test that indexes are correctly found. GEL tiering is the first in the
        list, so should be indexed at 0, and Exomiser is second, so should be
        indexed at 1
        '''
        self.genome_format = 'interpretedGenomes'
        self.genome_data_format = 'interpretedGenomeData'
        excel.index_interpretation_services(self)
        assert self.ex_index == 1 and self.gel_index == 0

    def test_error_raised_if_invalid_interpretation_service(self):
        '''
        Error should be raised if neither genomics_england_tiering' or
        'Exomiser' given as interpretation service
        '''
        self.genome_format = 'interpretedGenomes'
        self.genome_data_format = 'interpretedGenomeData'
        self.wgs_data["interpretedGenomes"][0]['interpretedGenomeData'][
                'interpretationService'
                ] = 'invalid_service'
        with pytest.raises(RuntimeError,
                    match="Interpretation services in JSON not recognised as "
                    "'genomics_england_tiering' or 'Exomiser'"):
            excel.index_interpretation_services(self)


class TestVariantInfo():

    @pytest.fixture
    def mock_variant(self):
        variant = {
                "coordinates": {
                    "chromosome": "12",
                    "start": 6936728,
                    "end": 6936773
                },
                "reportEvents": [
                    {
                        "tier": "TIER1",
                        "genomicEntities": [
                            {
                                "type": "gene",
                                "geneSymbol": "SYMB1"
                            }
                        ]
                    },
                    {
                        "tier": "TIER2",
                        "genomicEntities": [
                            {
                                "type": "gene",
                                "geneSymbol": "SYMB1"
                            }
                        ]
                    }
                ],
                "shortTandemRepeatReferenceData": {
                    "repeatedSequence": "CAG"
                },
                "variantCalls": [
                    {
                        "participantId": "testPB",
                        "numberOfCopies": [
                            {"numberOfCopies": 8},
                            {"numberOfCopies": 16}
                        ]
                    }
                ],
                "variantAttributes": {
                    "alleleFrequencies": ""
                }
            }
        return variant

    '''
    Test variant info functions.
    '''
    def test_add_cols_to_dict(self):
        '''
        Check that function takes list of columns and adds them to a dict with
        empty strings as the values.
        '''
        column_list = ["ColA", "ColB"]
        assert var_info.add_columns_to_dict(
            column_list
        ) == {"ColA": '', "ColB": ''}

    def test_get_str_info_tier1(self, mock_variant):
    # Mock input data

        proband = "testPB"
        columns = ["Chr", "Pos", "End", "Length", "Type", "Priority", "Repeat", "STR1", "STR2", "Gene", "AF Max"]
        ev_idx = 0
        proband_sex = "MALE"

        # Expected output for TIER1 STR
        expected_output = var_info.add_columns_to_dict(columns)
        expected_output.update({
            "Chr": "12",
            "Pos": 6936728,
            "End": 6936773,
            "Length": 45,
            "Type": "STR",
            "Priority": "TIER1_STR",
            "Repeat": "CAG",
            "STR1": 8,
            "STR2": 16,
            "Gene": "SYMB1",
            "AF Max": "-"
        })

        # Call the function to test for TEIR1
        result = var_info.get_str_info(mock_variant, proband, columns, ev_idx, proband_sex)

        # Assertions
        assert result == expected_output

    def test_get_str_info_tier2(self, mock_variant):
         # Modify the variant to replace reportEvents with TIER2 events
        mock_variant["reportEvents"] = [
            {
            "tier": "TIER2",
            "genomicEntities": [
                {
                "type": "gene",
                "geneSymbol": "SYMB1"
                }
            ]
            }
        ]

        proband = "testPB"
        columns = ["Chr", "Pos", "End", "Length", "Type", "Priority", "Repeat", "STR1", "STR2", "Gene", "AF Max"]
        ev_idx = 0
        proband_sex = "FEMALE"

        # Call the function to test for TIER2
        result = var_info.get_str_info(mock_variant, proband, columns, ev_idx, proband_sex)

        # Expected output for TIER2 STR
        expected_output_tier = var_info.add_columns_to_dict(columns)
        expected_output_tier.update({
            "Chr": "12",
            "Pos": 6936728,
            "End": 6936773,
            "Length": 45,
            "Type": "STR",
            "Priority": "TIER2_STR",
            "Repeat": "CAG",
            "STR1": 8,
            "STR2": 16,
            "Gene": "SYMB1",
            "AF Max": "-"
        })

        assert result == expected_output_tier

    def test_get_str_info_tier_null(self, mock_variant):
        mock_variant["reportEvents"] = [
            {
            "tier": "null",
            "genomicEntities": [
                {
                "type": "gene",
                "geneSymbol": "SYMB1"
                }
            ]
            }
        ]

        proband = "testPB"
        columns = ["Chr", "Pos", "End", "Length", "Type", "Priority", "Repeat", "STR1", "STR2", "Gene", "AF Max"]
        ev_idx = 0
        proband_sex = "MALE"

        # Call the function to test for null tier
        result = var_info.get_str_info(mock_variant, proband, columns, ev_idx, proband_sex)

        # Expected output for null tier STR
        expected_output_tier = var_info.add_columns_to_dict(columns)
        expected_output_tier.update({
            "Chr": "12",
            "Pos": 6936728,
            "End": 6936773,
            "Length": 45,
            "Type": "STR",
            "Priority": "null",
            "Repeat": "CAG",
            "STR1": 8,
            "STR2": 16,
            "Gene": "SYMB1",
            "AF Max": "-"
        })

        assert result == expected_output_tier

    def test_get_str_info_hemizygous(self, mock_variant):
        '''
        Check that the function returns the expected output in the case of
        missing X STR count in XY proband.
        '''

        mock_variant["coordinates"] = {
                "chromosome": "X",
                "start": 6936728,
                "end": 6936773
            }

        mock_variant["reportEvents"] = [
            {
            "tier": "TIER1",
            "genomicEntities": [
                {
                "type": "gene",
                "geneSymbol": "SYMB1"
                }
            ]
            }
        ]

        mock_variant["variantCalls"] = [
                {
                    "participantId": "testPB",
                    "numberOfCopies": [
                        {"numberOfCopies": 8}
                    ]
                },
                {
                    "participantId": "testT2",
                    "numberOfCopies": [
                        {"numberOfCopies": 14},
                        {"numberOfCopies": 16}
                    ]
                },
                {
                    "participantId": "testT3",
                    "numberOfCopies": [
                        {"numberOfCopies": 8}
                    ]
                }
            ]

        proband = "testPB"
        columns = ["Chr", "Pos", "End", "Length", "Type", "Priority", "Repeat", "STR1", "STR2", "Gene", "AF Max"]
        ev_idx = 0
        proband_sex = "MALE"

        # Call the function to test with hemizygous proband
        result = var_info.get_str_info(mock_variant, proband, columns, ev_idx, proband_sex)

        # Expected output for hemizygous proband
        expected_output = var_info.add_columns_to_dict(columns)
        expected_output.update({
            "Chr": "X",
            "Pos": 6936728,
            "End": 6936773,
            "Length": 45,
            "Type": "STR",
            "Priority": "TIER1_STR",
            "Repeat": "CAG",
            "STR1": 8,
            "STR2": "",
            "Gene": "SYMB1",
            "AF Max": "-"
        })

        assert result == expected_output


    def test_tier_conversion(self):
        '''
        Test Tiers from JSON are converted into tier representation as desired
        by workbook. Workbook tiers should include the tier and the variant
        type
        '''
        tiers_to_convert = [
            ["TIER1", "SNV"],
            ["TIER2", "SNV"],
            ["TIER1", "CNV"],
            ["TIERA", "CNV"],
            ["TIER1", "STR"],
            ["TIER2", "STR"],
        ]

        tiers = []
        for tiering in tiers_to_convert:
            tiers.append(var_info.convert_tier(tiering[0], tiering[1]))

        assert tiers == [
            "TIER1_SNV", "TIER2_SNV", "TIER1_CNV", "TIER1_CNV", "TIER1_STR", "TIER2_STR"
        ]

    def test_get_af_max(self):
        '''
        Test that the highest AF is returned for the variant.
        '''
        variant = {
            'variantAttributes': {
            'alleleFrequencies': [
                {
                    'alternateFrequency': 0.00003
                },
                {
                    'alternateFrequency': 0.001
                }
            ]
        }}
        assert var_info.get_af_max(variant) == '0.001'

    def test_male_proband_X_SNV_is_hemizygous(self):
        '''
        Placeholder for testing male proband X SNV hemizygosity.
        '''

        heterozygous_variant = "heterozygous"
        alt_hom_variant = "alternate_homozygous"

        assert var_info.get_zygosity(heterozygous_variant, "MALE", 'X') == 'hemizygous'
        assert var_info.get_zygosity(alt_hom_variant, "MALE", 'X') == 'hemizygous'

        assert var_info.get_zygosity(heterozygous_variant, "MALE", '12') == 'heterozygous'
        assert var_info.get_zygosity(heterozygous_variant, "FEMALE", 'X') == 'heterozygous'


class TestIndexParticipant():
    '''
    Tests for get_variant_info.index_participant
    '''
    proband = "p123456789"
    variant = {
        'variantCalls': [
            {
                'participantId': 'pXXXXXXXXX'
            },
            {
                'participantId': 'p123456789'
            },
            {
                'participantId': 'pYYYYYYYYY'
            }
    ]}
    def test_index_if_proband(self):
        '''
        Check indexing of proband is worked out correctly; here the proband is
        the second in the list, so we expect index 1 to be returned.
        '''
        assert var_info.index_participant(self.variant, self.proband) == 1

    def test_index_if_proband_not_found(self):
        '''
        Check indexing of participant errors if proband cannot be found
        '''
        self.variant['variantCalls'].pop(1)

        with pytest.raises(RuntimeError):
            var_info.index_participant(self.variant, self.proband)

    def test_returns_none_if_no_idx_provided(self):
        '''
        Check if index is None (i.e. there is no mother and/or father) None
        is returned.
        '''
        assert var_info.index_participant(self.variant, None) is None


class TestRanking():
    '''
    Tests for ranking function
    '''
    ranks = [1, 2, 3, 3, 4]
    str_ranks = [f"Exomiser Rank {str(x)}" for x in ranks]
    df = pd.DataFrame({'Priority': str_ranks})
    print(df)

    def test_can_handle_two_bronze(self):
        '''
        Check indices both third ranked items are returned.
        '''
        correct_ranks = self.str_ranks[:-1]

        pd.testing.assert_frame_equal(
            var_info.get_top_3_ranked(self.df),
            pd.DataFrame({'Priority': correct_ranks})
        )

    def test_next_ranked_returned_if_no_items_at_rank(self):
        '''
        Check that indices for the third and forth ranked items are returned if
        there is no second ranked item
        '''
        self.str_ranks.pop(1)
        self.df = self.df.drop([1])

        pd.testing.assert_frame_equal(
            var_info.get_top_3_ranked(self.df).reset_index(drop=True),
            pd.DataFrame({'Priority': self.str_ranks}).reset_index(drop=True)
        )



class TestVariantNomenclature():
    '''
    Test variant nomenclature functions.
    '''
    def test_get_ensp(self):
        '''
        Check that get_ensp function returns ENSP protein ID in the same list
        item as the ENST transcript ID
        '''
        refseq_tsv = ["ENST0000033\tENSP0000044\tENSG00000022",
                           "ENST0000066\tENSP0000088\tENSG00000044"]
        assert var_info.look_up_id_in_refseq_mane_conversion_file(
            refseq_tsv, "ENST0000033", "ENSP"
        ) == "ENSP0000044"


class TestHpoTerms():
    '''
    Tests for HPO unknown filtering function when "termPresence" is "unknown"
    '''
    @mock.patch('obonet.read_obo')
    def test_hpo_unknown_filtering_in_get_hpo_terms(self, mock_obo):
        '''
        Test get_hpo_terms function filters out terms with "termPresence" of "unknown"
        '''
        # Mock obo data
        mock_graph = MagicMock()
        mock_graph.nodes = {
            "HP:0004322": {"name": "Short stature"},
            "HP:0001249": {"name": "Intellectual disability"}
        }
        mock_obo.return_value = mock_graph

        # Mock member data with mixed termPresence values
        member = {
            "hpoTermList": [
                {
                    "term": "HP:0004322",
                    "termPresence": "present"
                },
                {
                    "term": "HP:0001249",
                    "termPresence": "unknown"
                }
            ]
        }
        mock_args = MagicMock()
        excel_instance = excel(mock_args)
        result = excel_instance.get_hpo_terms(member)

        # Should only include the "present" only
        assert result == "Short stature"


class TestInterpretationFlags():
    '''
    Tests for interpretation flags extraction from JSON request
    '''
    wgs_data = {
        "family_id": "FAM12345",
        "interpretation_request_data": {
            "json_request": {
                "interpretation_flags": "Flag1, Flag2, Flag3"
            }
        },
    }
    summary_content = {}

    def test_interpretation_flags_extraction(self,wgs_data=wgs_data):
        '''
        Test that interpretation flags are correctly extracted from JSON request
        and added to summary_content dictionary.
        '''
        mock_args = MagicMock()
        excel_instance = excel(mock_args)
        result = excel_instance.get_summary_content(wgs_data)
        summary_content = result.get((1, 9))
        expected_flags = "Flag1, Flag2, Flag3"
        assert summary_content == expected_flags
        assert result[(1, 2)] == "FAM12345"


class TestWorkbookName:
    def test_workbook_name_generation(self):
        """
        Test that the workbook name is generated correctly based on family_id.
        """
        family_id = "FAM12345"
        expected_workbook_name = f"{family_id}.xlsx"

        mock_args = MagicMock()
        mock_args.output_filename = None
        mock_args.acmg = None
        mock_args.cnv = None

        excel_instance = excel(mock_args)
        excel_instance.wgs_data = {'family_id': family_id}

        with patch.object(excel_instance, 'open_files'), \
             patch('dxpy.find_data_objects', return_value=[]), \
             patch('pandas.ExcelWriter'), \
             patch.object(excel_instance, 'summary_page'), \
             patch.object(excel_instance, 'get_interpreted_genome_format'), \
             patch.object(excel_instance, 'index_interpretation_services'), \
             patch.object(excel_instance, 'create_gel_tiering_variant_page'), \
             patch.object(excel_instance, 'create_additional_analysis_page'), \
             patch.object(excel_instance, 'str_image_page'), \
             patch.object(excel_instance, 'writer', create=True), \
             patch.object(excel_instance, 'workbook', create=True), \
             patch('excel_styles.DropDown.drop_down'):

            excel_instance.generate()
        print("Generated filename:", excel_instance.args.output_filename)
        assert excel_instance.args.output_filename == expected_workbook_name


class TestExomiomiserDenovoDuplicates:
    @staticmethod
    def make_mock_variant(row, source='exomiser'):
        """
        Helper function to create a mock variant dictionary based on a DataFrame row.
        Designed for filtering tests — includes minimal required fields to avoid errors.
        """
        return {
            'variantCoordinates': {
                'chromosome': row['Chr'],
                'position': row['Pos'],
                'reference': row['Ref'],
                'alternate': row['Alt']
            },
            'variantCalls': {
                None: {
                    "zygosity": "HET",
                    "depthAlternate": 50
                }
            },
            'variantAttributes': {
                'alleleFrequencies': [],
                'additionalTextualVariantAnnotations': {
                    'hgvs': f"{row['Gene']}:ENST000001:c.123A>T:p.Lys41Asn"
                }
            },
            'reportEvents': [{
                'tier': '',
                'score': 0.8 if source == 'exomiser' else 0.9,
                'domain': '',
                'actions': None,
                'penetrance': '',
                'modeOfInheritance': '',
                'segregationPattern': 'de novo' if source == 'gel' else '',
                'genePanel': {},
                'haplotype': None,
                'phenotypes': {},
                'consequences': None,
                'roleInCancer': None,
                'evidenceEntry': {},
                'vendorSpecificScores': {
                    'rank': int(row['Priority'][-1]) if 'Rank' in row['Priority'] else 1
                },
                'gene': {'symbol': row['Gene']},
                'genomicEntities': [],
                'additionalTextualVariantAnnotations': {}
            }]
        }

    def test_remove_denovo_duplicates(self):
        exomiser_data = {
            'Chr': ['1', '1', '2'],
            'Pos': [100, 200, 300],
            'Ref': ['A', 'G', 'T'],
            'Alt': ['C', 'T', 'G'],
            'Priority': ['Exomiser Rank 1', 'Exomiser Rank 2', 'Exomiser Rank 3'],
            'Gene': ['GENE1', 'GENE2', 'GENE3']
        }
        denovo_data = {
            'Chr': ['1', '2', '3'],
            'Pos': [100, 400, 500],
            'Ref': ['A', 'C', 'A'],
            'Alt': ['C', 'A', 'G'],
            'Priority': ['de Novo', 'de Novo', 'de Novo'],
            'Gene': ['GENE1', 'GENE4', 'GENE5']
        }

        expected_filtered = pd.DataFrame({
            'Chr': ['1', '1', '2', '2', '3'],
            'Pos': [100, 200, 300, 400, 500],
            'Ref': ['A', 'G', 'T', 'C', 'A'],
            'Alt': ['C', 'T', 'G', 'A', 'G'],
            'Priority': ['Exomiser Rank 1', 'Exomiser Rank 2', 'Exomiser Rank 3', 'de Novo', 'de Novo'],
            'Gene': ['GENE1', 'GENE2', 'GENE3', 'GENE4', 'GENE5']
        })

        mock_args = MagicMock()
        excel_instance = excel(mock_args)

        excel_instance.genome_format = 'GRCh38'
        excel_instance.ex_index = 'exomiser'
        excel_instance.gel_index = 'gel'
        excel_instance.genome_data_format = 'genome'
        excel_instance.proband = None
        excel_instance.mother = None
        excel_instance.father = None
        excel_instance.proband_sex = None
        excel_instance.column_list = expected_filtered.columns.tolist()
        excel_instance.mane = []
        excel_instance.refseq_tsv = []
        excel_instance.var_df = pd.DataFrame()

        original_exomiser_df = pd.DataFrame(exomiser_data)
        denovo_df = pd.DataFrame(denovo_data)

        print("\nOriginal Exomiser DataFrame:")
        print(original_exomiser_df)

        print("\nOriginal De Novo DataFrame:")
        print(denovo_df)

        excel_instance.wgs_data = {
            'GRCh38': {
                'exomiser': {
                    'genome': {
                        'variants': [
                            self.make_mock_variant(row, source='exomiser')
                            for _, row in original_exomiser_df.iterrows()
                        ]
                    }
                },
                'gel': {
                    'genome': {
                        'variants': [
                            self.make_mock_variant(row, source='gel')
                            for _, row in denovo_df.iterrows()
                        ]
                    }
                }
            }
        }

        captured = {}

        def test_capture(self_df, **kwargs):
            # Only capture the Extended_analysis sheet
            if kwargs.get('sheet_name') == 'Extended_analysis':
                if isinstance(self_df, pd.DataFrame):
                    captured['Extended_analysis'] = self_df.copy()
                elif isinstance(self_df, MagicMock):
                    captured['Extended_analysis'] = expected_filtered.copy()
                else:
                    raise ValueError("Expected pandas df")

        with patch.object(excel_instance, 'open_files'), \
            patch.object(pd.DataFrame, 'to_excel') as to_excel_mock, \
            patch.object(excel_instance, 'writer', create=True), \
            patch.object(excel_instance, 'workbook', create=True):

            to_excel_mock.side_effect = test_capture
            excel_instance.writer = MagicMock()

            # Run the method
            excel_instance.create_additional_analysis_page()

            # Extract and compare
            actual_df = captured['Extended_analysis'][['Chr','Pos','Ref','Alt','Priority','Gene']]\
                .sort_values(by=['Chr', 'Pos']).reset_index(drop=True)
            expected_df = expected_filtered.sort_values(by=['Chr', 'Pos']).reset_index(drop=True)

            pd.testing.assert_frame_equal(actual_df, expected_df)

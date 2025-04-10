#!/usr/bin/env python3
import argparse
import pandas as pd
import pytest
import os
import obonet
import sys
import json
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
        "family_id": "r12345",
        "interpretation_request_data": {
            "json_request": {
                "pedigree": {
                    "members": [
                        {
                            "hpoTermList": [
                                {"hpoBuildNumber": "vXXXXXX"}
                            ]
                        }
                    ],
                    'diseasePenetrances': [
                        {
                            'penetrance': 'complete',
                            'specificDisease': 'Congenital malformation'
                        },
                        {
                            'penetrance': 'incomplete',
                            'specificDisease': 'OtherDisease'                   
                        }
                    ],
                    'analysisPanels': [
                        {
                            'panelId': "486",
                            'panelName': "286",
                            'specificDisease': 'Congenital malformation',
                            'panelVersion': "2.2"
                        }
                    ]
                }
            }
        }
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

    def test_get_str_info_tier1(self, variant=variant):
    # Mock input data
        
        proband = "testPB"
        columns = ["Chr", "Pos", "End", "Length", "Type", "Priority", "Repeat", "STR1", "STR2", "Gene", "AF Max"]
        ev_idx = 0

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
        result = var_info.get_str_info(variant, proband, columns, ev_idx)

        # Assertions
        assert result == expected_output

    def test_get_str_info_tier2(self, variant=variant):
         # Modify the variant to replace reportEvents with TIER2 events
        variant["reportEvents"] = [
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

        # Call the function to test for TIER2
        result = var_info.get_str_info(variant, proband, columns, ev_idx)

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

    def test_get_str_info_tier_null(self, variant=variant):
        variant["reportEvents"] = [
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

        # Call the function to test for TIER2
        result = var_info.get_str_info(variant, proband, columns, ev_idx)

        # Expected output for TIER2 STR
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

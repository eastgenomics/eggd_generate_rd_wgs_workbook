#!/usr/bin/env python3
import argparse
import pytest
import os
import obonet
import sys
import json
from make_workbook import excel
from get_variant_info import VariantNomenclature, VariantUtils
from start_process import SortArgs
from unittest import mock
from unittest.mock import MagicMock, patch


class TestWorkbook():
    '''
    Tests for excel() class in make_workbook script
    '''
    wgs_data = {
        "referral": {
            "referral_data": {
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
                            'specificDisease': 'Disease'
                        },
                        {
                            'penetrance': 'incomplete',
                            'specificDisease': 'OtherDisease'                   
                        }
                    ]
                },
                'referralTests': [{
                'analysisPanels': [
                        {'panelId': "486",
                        'panelName': "286",
                        'specificDisease': 'Disease',
                        'panelVersion': "2.2"}
                        ]
                    }
                ]
            }
        }
    }

    def test_get_panels(self):
        '''
        Check that panels are extracted from JSON as expected.
        '''
        self.summary_content = {}
        excel.get_panels(self)
        assert self.summary_content == {
            (14, 1): '486', (14, 2): 'Disease', (14, 3): '2.2', (14, 4): '286'
        }

    def test_get_penetrance(self):
        '''
        Check that penetrance is extracted from JSON as expected, and matched
        to the specific disease in the referral
        '''
        self.summary_content = {(2,2): 'Disease'}
        excel.get_penetrance(self)
        assert self.summary_content[(3,2)] == "complete"


class TestInterpretationService():
    '''
    Test that the function to find interpretation service works as expected
    '''
    wgs_data = {
        'interpretedGenomes': [
            {'interpretedGenomeData': {
                'interpretationService': 'genomics_england_tiering'}
            },
            {'interpretedGenomeData': {'interpretationService': 'Exomiser'}}
        ]
    }
    def test_indexing_of_interpretation_service(self):
        '''
        Test that indexes are correctly found. GEL tiering is the first in the
        list, so should be indexed at 0, and Exomiser is second, so should be
        indexed at 1
        '''
        excel.index_interpretation_services(self)
        assert self.ex_index == 1 and self.gel_index == 0

    def test_error_raised_if_invalid_interpretation_service(self):
        '''
        Error should be raised if neither genomics_england_tiering' or
        'Exomiser' given as interpretation service
        '''
        self.wgs_data["interpretedGenomes"][0]['interpretedGenomeData'][
                'interpretationService'
                ] = 'invalid_service'
        with pytest.raises(RuntimeError,
                    match="Interpretation services in JSON not recognised as "
                    "'genomics_england_tiering' or 'Exomiser'"):
            excel.index_interpretation_services(self)


class TestVariantInfo():
    '''
    Test variant info functions.
    '''
    def test_add_cols_to_dict(self):
        '''
        Check that function takes list of columns and adds them to a dict with
        empty strings as the values.
        '''
        column_list = ["ColA", "ColB"]
        assert VariantUtils.add_columns_to_dict(
            column_list
        ) == {"ColA": '', "ColB": ''}

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
        ]

        tiers = []
        for tiering in tiers_to_convert:
            tiers.append(VariantUtils.convert_tier(tiering[0], tiering[1]))

        assert tiers == [
            "TIER1_SNV", "TIER2_SNV", "TIER1_CNV", "TIER1_CNV", "TIER1_STR"
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
        assert VariantUtils.get_af_max(variant) == 0.001


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
        assert VariantUtils.index_participant(self.variant, self.proband) == 1

    def test_index_if_proband_not_found(self):
        '''
        Check indexing of participant errors if proband cannot be found
        '''
        self.variant['variantCalls'].pop(1)

        with pytest.raises(RuntimeError):
            VariantUtils.index_participant(self.variant, self.proband)

    def test_returns_none_if_no_idx_provided(self):
        '''
        Check if index is None (i.e. there is no mother and/or father) None
        is returned.
        '''
        assert VariantUtils.index_participant(self.variant, None) is None


class TestRanking():
    '''
    Tests for ranking function
    '''
    snvs = [
        {'reportEvents': {'vendorSpecificScores': {'rank': 1}}},
        {'reportEvents': {'vendorSpecificScores': {'rank': 2}}},
        {'reportEvents': {'vendorSpecificScores': {'rank': 3}}},
        {'reportEvents': {'vendorSpecificScores': {'rank': 3}}},
        {'reportEvents': {'vendorSpecificScores': {'rank': 4}}}
    ]

    def test_can_handle_two_bronze(self):
        '''
        Check both third ranked items are returned.
        '''
        assert VariantUtils.get_top_3_ranked(self.snvs) == [
            {'reportEvents': {'vendorSpecificScores': {'rank': 1}}},
            {'reportEvents': {'vendorSpecificScores': {'rank': 2}}},
            {'reportEvents': {'vendorSpecificScores': {'rank': 3}}},
            {'reportEvents': {'vendorSpecificScores': {'rank': 3}}}
        ]

    def test_next_ranked_returned_if_no_items_at_rank(self):
        '''
        Check that third and forth ranked items are returned if there is no
        second ranked item
        '''
        self.snvs[1] = {'reportEvents': {'vendorSpecificScores': {'rank': 3}}}
        assert VariantUtils.get_top_3_ranked(self.snvs) == [
            {'reportEvents': {'vendorSpecificScores': {'rank': 1}}},
            {'reportEvents': {'vendorSpecificScores': {'rank': 3}}},
            {'reportEvents': {'vendorSpecificScores': {'rank': 3}}},
            {'reportEvents': {'vendorSpecificScores': {'rank': 3}}},
            {'reportEvents': {'vendorSpecificScores': {'rank': 4}}}
        ]


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
        assert VariantNomenclature.get_ensp(
            refseq_tsv, "ENST0000033"
        ) == "ENSP0000044"

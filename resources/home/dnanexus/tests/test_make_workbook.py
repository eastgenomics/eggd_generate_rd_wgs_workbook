#!/usr/bin/env python3
import argparse
import pytest
import os
import sys
from make_workbook import excel
from get_variant_info import VariantNomenclature, VariantInfo
from start_process import SortArgs
from unittest import mock
from unittest.mock import MagicMock, patch

class TestWorkbook():
    '''
    Tests for excel() class in make_workbook script
    '''

    @mock.patch('argparse.ArgumentParser.parse_args',
            return_value=argparse.Namespace(obo_files=True))
    def test_invalid_hpo_version_dx(self, mock):
        '''
        Test that if the HPO version is invalid a RunTime error is passed for
        obo file arrays on DNAnexus
        '''
        self.args = mock
        self.wgs_data = {'referral': {'referral_data': {'pedigree': {'members': [{
            'hpoTermList': [
                {'hpoBuildNumber': "vXXXXXX"}
            ]
        }]}}}}

        with pytest.raises(RuntimeError):
            excel.get_hpo_obo(self)

    @mock.patch('argparse.ArgumentParser.parse_args',
            return_value=argparse.Namespace(
                obo_path='/path/to/wherever/', obo_files=None
                ))
    def test_invalid_hpo_versions_path(self, mock):
        '''
        Test that if the HPO version is invalid a RunTime error is passed for
        obo paths run locally
        '''
        self.wgs_data = {'referral': {'referral_data': {'pedigree': {'members': [{
            'hpoTermList': [
                {'hpoBuildNumber': "vXXXXXX"}
            ]
        }]}}}}
        self.args = mock
        with pytest.raises(RuntimeError):
            excel.get_hpo_obo(self)
    
    @mock.patch('argparse.ArgumentParser.parse_args',
            return_value=argparse.Namespace(obo_path=None, obo_files=True))
    def test_correct_hpo_version_dx(self, mock):
        '''
        Test that the correct path to obo is given based on the version
        specified. This test is for obo_files input from DNAnexus
        '''
        self.wgs_data = {'referral': {'referral_data': {'pedigree': {'members': [{
            'hpoTermList': [
                {'hpoBuildNumber': "v2019_02_12"}
            ]
        }]}}}}
        self.args = mock
        assert excel.get_hpo_obo(self) == "/home/dnanexus/obo_files/hpo_v20190212.obo"

    # @mock.patch('argparse.ArgumentParser.parse_args',
    #         return_value=argparse.Namespace(
    #             obo_path='/path/to/wherever/', obo_files=None
    #             ))
    # def test_correct_hpo_version_path(self, mock):
    #     '''
    #     Test that the correct path to obo is given based on the version
    #     specified. This test is for obo_path input from the command line
    #     '''
    #     self.wgs_data = {'referral': {'referral_data': {'pedigree': {'members': [{
    #         'hpoTermList': [
    #             {'hpoBuildNumber': "v2019_02_12"}
    #         ]
    #     }]}}}}
    #     self.args = mock
    #     assert excel.get_hpo_obo(self) == "/path/to/wherever/hpo_v20190212.obo"
    
    def test_get_panels(self):
        '''
        Check that panels are extracted from JSON as expected.
        '''
        self.summary_content = {}
        self.wgs_data = {'referral': {'referral_data': {'referralTests': [{
            'analysisPanels': [
                {'panelId': "486",
                 'panelName': "286",
                 'specificDisease': 'Disease',
                 'panelVersion': "2.2"}
            ]
        }]}}}
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
        self.wgs_data = {'referral': {'referral_data': {'pedigree':
            {'diseasePenetrances':
                [{'penetrance': 'complete',
                    'specificDisease': 'Disease'
                },
                {
                    'penetrance': 'incomplete',
                    'specificDisease': 'OtherDisease'                   
                }
                ]}}}}
        excel.get_penetrance(self)
        assert self.summary_content[(3,2)] == "complete"

    # def test_person_data(self):
    #     self.wgs_data = {'referral': {'referral_data': {'pedigree':
    #         {'members':
    #          {
    #              "isProband": 'True'
    #          }

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
        assert VariantInfo.add_columns_to_dict(column_list) == {"ColA": '', "ColB": ''}

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
            tiers.append(VariantInfo.convert_tier(tiering[0], tiering[1]))
        
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
        assert VariantInfo.get_af_max(variant) == 0.001


class TestCheckIfProband():
    '''
    Tests for excel.check_if_proband function
    '''
    proband = "p123456789"
    variantCalls = [
        {
            'participantId': 'pXXXXXXXXX'
        },
        {
            'participantId': 'p123456789'
        },
        {
            'participantId': 'pYYYYYYYYY'
        }
    ]
    def test_check_if_proband(self):
        '''
        Check indexing of proband is worked out correctly; here the proband is
        the second in the list, so we expect index 1 to be returned.
        '''
        assert excel.check_if_proband(self, self.variantCalls) == 1

    def test_check_if_proband_not_found(self):
        '''
        Check indexing of proband errors if proband cannot be found
        '''
        self.variantCalls.pop(1)

        with pytest.raises(RuntimeError):
            excel.check_if_proband(self, self.variantCalls)


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

    
    # @patch('excel.str_image_page')
    # def test_str_image_page(self, openpyxl_mock):
    #     self.workbook = openpyxl_mock.workbook()
    #     self.assertEqual(
    #         excel.str_image_page(
    #             sheetname="STR guidelines",
    #             cell='B2'
    #         )
    #     ), (
    #         "From CU-WG-REF-40 Guidelines for Rare Disease Whole Genome "
    #         "Sequencing & Next Generation Sequencing Panel Interpretation & "
    #         "Reporting"
    #     )


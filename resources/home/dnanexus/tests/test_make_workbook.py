#!/usr/bin/env python3
import argparse
import pytest
import os
import sys
from make_workbook import excel
from start_process import SortArgs
from unittest import mock

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
        version = "vXXXXX"
        self.args.obo_files = True
        with pytest.raises(RuntimeError):
            excel.hpo_version(self, version)

    @mock.patch('argparse.ArgumentParser.parse_args',
            return_value=argparse.Namespace(
                obo_path='/path/to/wherever/', obo_files=None
                ))
    def test_invalid_hpo_versions_path(self, mock):
        '''
        Test that if the HPO version is invalid a RunTime error is passed for
        obo paths run locally
        '''
        version = "vXXXXX"
        with pytest.raises(RuntimeError):
            excel.hpo_version(mock, version)
    
    @mock.patch('argparse.ArgumentParser.parse_args',
            return_value=argparse.Namespace(obo_path=None, obo_files=True))
    def test_correct_hpo_version_dx(self, mock):
        '''
        Test that the correct path to obo is given based on the version
        specified. This test is for obo_files input from DNAnexus
        '''
        version = "v2019_02_12"
        assert excel.hpo_version(
            mock, version
            ) == "/home/dnanexus/obo_files/hpo_v20190212.obo"

    # @mock.patch('argparse.ArgumentParser.parse_args',
    #         return_value=argparse.Namespace(
    #             obo_path='/path/to/wherever/'
    #             ))
    # def test_correct_hpo_version_path(self, mock):
    #     '''
    #     Test that the correct path to obo is given based on the version
    #     specified. This test is for obo_files input from DNAnexus
    #     '''
    #     #self.args = SortArgs.parse_args(self)
    #     self.args = mock
    #     print(mock.obo_files)
    #     version = "v2019_02_12"
    #     assert excel.hpo_version(
    #         self, version
    #         ) == "/path/to/wherever/hpo_v20190212.obo"
    
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
            tiers.append(excel.convert_tier(self, tiering[0], tiering[1]))
        
        assert tiers == [
            "TIER1_SNV", "TIER2_SNV", "TIER1_CNV", "TIER1_CNV", "TIER1_STR"
        ]

    def test_add_cols_to_dict(self):
        '''
        Check that function takes list of columns and adds them to a dict with
        empty strings as the values.
        '''
        self.column_list = ["ColA", "ColB"]
        assert excel.add_columns_to_dict(self) == {"ColA": '', "ColB": ''}



        
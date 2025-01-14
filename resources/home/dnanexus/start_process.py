from make_workbook import excel
import dxpy
import json
import argparse
import os
from pathlib import Path


class SortArgs():
    '''
    Parse command line input/
    '''
    def __init__(self):
        self.args = self.parse_args()

    def parse_args(self):
        '''
        Parse command line arguments
        '''
        parser = argparse.ArgumentParser()
        parser.add_argument(
            "--json",
            help=(
                "WGS JSON from GEL to convert into workbook"
            )
        )
        parser.add_argument(
            '-o', '--output_filename', required=False,
            help=(
                'output name prefix for file, if not set will use json input'
            )
        )
        parser.add_argument(
            '--out_dir', required=False, default=os.getcwd(),
            help="path to output directory"
        )
        parser.add_argument(
            '--acmg', type=int, default=3,
            help='add extra ACMG reporting template sheet(s)'
        )
        parser.add_argument(
            '--cnv', type=int, default=2,
            help='add CNV reporting template sheet(s)'
        )
        parser.add_argument(
            '--mane_file',
            help='MANE list from Ensembl, mapping Ensembl IDs to RefSeq IDs'
        )
        parser.add_argument(
            '--config',
            help='Config file with HPO versions and thresholds for de novo '
            'quality score'
        )
        parser.add_argument(
            '--refseq_tsv',
            help='Refseq tsv, used to get Ensembl protein IDs'
        )
        parser.add_argument(
            '--epic_clarity',
            help='Epic Clarity export'
        )
        parser.add_argument(
            '--panels',
            help='JSON file mapping panel ID to panel name and test code'
        )
        return parser.parse_args()


def main():
    '''
    Pass args to excel() class
    '''
    parser = SortArgs()

    excel_handler = excel(
        args=parser.args
    )

    excel_handler.generate()


if __name__ == "__main__":
    main()

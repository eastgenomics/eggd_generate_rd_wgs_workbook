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
        self.parse_output()

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
            '-o', '--output', required=False,
            help=(
                'output name prefix for file, if not set will use json input'
            )
        )
        parser.add_argument(
            '--out_dir', required=False, default=os.getcwd(),
            help="path to output directory"
        )
        parser.add_argument(
            '--acmg', type=int, default=2,
            help='add extra ACMG reporting template sheet(s)'
        )
        parser.add_argument(
            '--cnv', type=int, default=2,
            help='add CNV reporting template sheet(s)'
        )
        parser.add_argument(
            '--mane',
            help='MANE list from Ensembl, mapping Ensembl IDs to RefSeq IDs'
        )
        parser.add_argument(
            '--obo_path',
            help='Path to Human Phenotype Ontology .obo files'
        )
        parser.add_argument(
            '--obo_files',
            help='Array of Human Phenotype Ontology .obo files'
        )
        parser.add_argument(
            '--refseq_tsv',
            help='Refseq tsv, used to get Ensembl protein IDs'
        )
        return parser.parse_args()

    def parse_output(self) -> None:
        """
        Strip JSON suffix, then set output to include outdir for writing output
        file. Note for DNAnexus files this will be overwritten later as it
        currently uses the DX file ID.
        """
        if not self.args.output:
            
            self.args.output = Path(
                self.args.json).name.replace('.json', '')

        self.args.output = (
            f"{Path(self.args.out_dir)}/{self.args.output}.xlsx"
        )


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

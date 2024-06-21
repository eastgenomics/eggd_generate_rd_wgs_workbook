#!/bin/bash
set -exo pipefail

main() {
    # Set up and download inputs
    mkdir -p /home/dnanexus/out/xlsx_reports
    dx-download-all-inputs --parallel
    # Add pip install location to path
    export PATH=$PATH:/home/dnanexus/.local/bin
    # Install packages
    sudo -H python3 -m pip install --no-index --no-deps packages/*

    # Build arguments
    args=""
    if [ "$acmg" ]; then args+="--acmg ${acmg} "; fi
    if [ "$cnv" ]; then args+="--cnv ${cnv} "; fi
    if [ "$output_filename" ]; then args+="--output_file_name ${output_filename} "; fi
    if [ "$epic_clarity" ]; then args+="--epic_clarity /home/dnanexus/in/epic_clarity/*.xlsx "; fi

    # Generate workbook
    /usr/bin/time -v python3 start_process.py \
    --json /home/dnanexus/in/json/*json \
    --mane_file /home/dnanexus/in/mane_file/* \
    --refseq_tsv /home/dnanexus/in/refseq_tsv/*tsv \
    --config /home/dnanexus/in/config/*json \
    $args

    mv *.xlsx /home/dnanexus/out/xlsx_reports
    # Upload workbook
    output_xlsx=$(dx upload /home/dnanexus/out/xlsx_reports/* --brief)
    dx-jobutil-add-output xlsx_report "$output_xlsx" --class=file
}
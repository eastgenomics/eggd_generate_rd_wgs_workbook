#!/bin/bash
set -exo pipefail

main() {
    #mark-section "Setting up and downloading inputs"
    mkdir -p /home/dnanexus/out/xlsx_reports
    dx-download-all-inputs --parallel
    # add pip install location to path
    export PATH=$PATH:/home/dnanexus/.local/bin
    # install packages
    sudo -H python3 -m pip install --no-index --no-deps packages/*

    #mark-section "Building arguments"
    args=""
    if [ "$acmg" ]; then args+="--acmg ${acmg} "; fi
    if [ "$cnv" ]; then args+="--cnv ${cnv} "; fi
    if [ "$output_name_prefix" ]; then args+="--output ${output} "; fi
    if [ "$out_dir" ]; then args+="--out_dir ${out_dir} "; fi
    if [ "$acmg" ]; then args+="--acmg ${acmg} "; fi
    if [ "$obo_path" ]; then args+="--obo_path ${obo_path} "; fi
    if [ "$obo_files" ]; then args+="--obo_files True "; fi

    ls ~/in/obo_files/0
    ls ~/in/obo_files/1
    mkdir ~/obo_files

    # Download obo files into single folder
    if [ "$obo_files" ]; then find ~/in/obo_files -type f -name "*" -print0 | xargs -0 -I {} mv {} ~/obo_files; fi

    #mark-section "Generating workbook"
    /usr/bin/time -v python3 start_process.py \
    --json /home/dnanexus/in/json/*json --mane_file /home/dnanexus/in/mane/* --refseq_tsv /home/dnanexus/in/refseq_tsv/*tsv $args
    mv *.xlsx /home/dnanexus/out/xlsx_reports
    #mark-section "Uploading workbook"
    output_xlsx=$(dx upload /home/dnanexus/out/xlsx_reports/* --brief)
    dx-jobutil-add-output xlsx_report "$output_xlsx" --class=file
}
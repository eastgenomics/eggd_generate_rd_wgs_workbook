import re


def add_columns_to_dict(column_list):
    '''
    Function to add columns to variant pages.
    All columns are empty strings, so they can be overwritten by values if
    needed, but left blank if not.
    Inputs:
        column_list (list): a list of column names
    Outputs:
        variant_dict (dict): a dictionary with each item in column_list as
        a key, and an empty string as the value
    '''
    variant_dict = {}
    for col in column_list:
        variant_dict[col] = ""

    return variant_dict


def get_gene_symbol(variant):
    '''
    Get gene symbol from variant record. Some records will have more than
    one gene symbol; in those cases, all gene symbols will be returned.
    Inputs:
        variant (dict): record for that specific variant from the JSON
    Outputs:
        gene_symbol (str): the gene symbol for that variant
    '''
    gene_list = []
    for entry in variant['reportEvents'][0]['genomicEntities']:
        if entry['type'] == 'gene':
            gene_list.append(entry['geneSymbol'])
    uniq_genes = sorted(list(set(gene_list)))
    if len(uniq_genes) == 1:
        gene_symbol = uniq_genes.pop()
    else:
        gene_symbol = str(uniq_genes).strip('[').strip(']')
    return gene_symbol


def convert_tier(tier, var_type):
    '''
    Convert tier from GEL tiering variant to include variant type. This is
    to facilitate counting the total variants of each tier and type, as
    well as to convert tier "A" to "1" for consistency.
    Inputs:
        tier (str): described GEL tiering tier for a variant in the JSON
        var_type (str): variant type
    Outputs:
        tier (str): tier for that variant converted to include variant type
        to facilitate counting a total number of variants for each tier and
        type
    '''
    if tier == "TIER1" and var_type == "SNV":
        tier = "TIER1_SNV"
    elif tier == "TIER1" and var_type == "CNV":
        tier = "TIER1_CNV"
    elif tier == "TIER2" and var_type == "SNV":
        tier = "TIER2_SNV"
    elif tier == "TIER2" and var_type == "CNV":
        tier = "TIER2_CNV"
    elif tier == "TIERA" and var_type == "CNV":
        tier = "TIER1_CNV"
    elif tier == "TIER1" and var_type == "STR":
        tier = "TIER1_STR"
    elif tier == "TIER2" and var_type == "STR":
        tier = "TIER2_STR"
    return tier


def get_af_max(variant):
    '''
    Get AF for population with highest allele frequency in the JSON
    Inputs:
        variant (dict): dict describing a single variant from JSON
    Outputs:
        highest_af: frequency of the variant in the population with
        the highest allele frequency for the variant, or - if the variant
        is not seen in any populations in the JSON, or - if there are no
        reference populations.
    '''
    highest_af = 0
    if variant['variantAttributes']['alleleFrequencies'] is None:
        highest_af = '-'
    else:
        for af in variant['variantAttributes']['alleleFrequencies']:
            if af['alternateFrequency'] > highest_af:
                highest_af = af['alternateFrequency']
        if highest_af == 0:
            highest_af = '-'

    return str(highest_af)


def get_inheritance(variant, mother_idx, father_idx, p_sex):
    '''
    Work out inheritance of variant based on zygosity of parents
    Inputs:
        variant: (dict): dict extracted from JSON describing single variant
        mother_idx: (int): position within list of variant calls that
        belongs to the mother, or None if no mother in JSON
        father_idx: (int): position within list of variant calls that
        belongs to the father, or None if no father in JSON
        p_sex (str): proband sex
    Outputs:
        inheritance (str): inferred inheritance of the variant, or None.
    '''
    zygosity = lambda x, y: x['variantCalls'][y]['zygosity']

    inheritance = None
    maternally_inherited = False
    paternally_inherited = False

    inheritance_types = ['alternate_homozygous', 'heterozygous', 'hemizygous']

    # if there is a mother in the JSON and the variant is alt_homozygous in
    # or heterozygous in the mother then can infer maternal inheritance
    if (mother_idx is not None and
        zygosity(variant, mother_idx) in inheritance_types):
        maternally_inherited = True

    # if there is a father in the JSON and the variant is alt_homozygous in
    # or heterozygous in the father then can infer paternal inheritance
    # filter out XY probands here as X should be inherited from mother
    if (father_idx is not None and
        zygosity(variant, father_idx) in inheritance_types and
        not (p_sex == 'MALE' and
        variant['variantCoordinates']['chromosome'] == 'X')):
        paternally_inherited = True

    if maternally_inherited and not paternally_inherited:
        inheritance = "maternal"

    elif paternally_inherited and not maternally_inherited:
        inheritance = "paternal"

    elif maternally_inherited and paternally_inherited:
        inheritance = "both"

    return inheritance


def convert_moi(moi):
    '''
    Convert MOI for SNV to use more human readable wording
    Inputs
        moi (str): ModeOfInheritance field from GEL JSON
    Outputs
        Converted ModeOfInheritance
    '''
    conversion = {
        "biallelic": "Autosomal Recessive",
        "monoallelic_not_imprinted": "Autosomal Dominant",
        "monoallelic_paternally_imprinted": "Autosomal Dominant - "
        "Paternally Imprinted",
        "monoallelic_maternally_imprinted": "Autosomal Dominant - "
        "Maternally Imprinted",
        "xlinked_biallelic": "X-Linked Recessive",
        "xlinked_monoallelic": "X-Linked Dominant",
        "mitochondrial": "Mitochondrial"
    }

    try:
        c_moi = conversion[moi]
    except KeyError:
        c_moi = None
        print(
            f"Could not find ModeOfInheritance {moi} in conversion table"
            f"{conversion}. Setting MOI to None for this variant."
        )
    return c_moi


def index_participant(variant, participant_id):
    '''
    Take list of variantCalls for a variant and return index of the
    participant.
    Inputs:
        variant: (dict): dict extracted from JSON describing single variant
        participant_id (str): GEL ID for the participant
    Outputs:
        index (int): index of variantCalls list in dict for the participant
    '''
    index = None
    if participant_id is not None:
        for call in variant['variantCalls']:
            if call['participantId'] == participant_id:
                index = variant['variantCalls'].index(call)
                break

        if index is None:
            raise RuntimeError(
                f"Unable to find participant ID {participant_id} in "
                f"variant {variant['variantCalls']}"
            )

    return index


def get_str_info(variant, proband, columns):
    '''
    Each variant that will be added to the excel workbook, needs to be
    added to the dataframe via a dictionary of values for each column
    heading in the workbook

    This function creates a variant dict for specific STR variant to be
    added to the excel workbook, where the keys are the columns in the
    workbook and the values are the values for this variant.
    Inputs:
        variant (dict): dict extracted from JSON describing single variant
        proband (str): GEL ID for the proband
        columns (list): list of columns to make into keys for variant dict
    Outputs:
        var_dict: (dict) dict of variant information extracted from JSON,
        formatted with the correct column headings for the excel workbook.
        this will be added to a list of dicts for conversion into dataframe
    '''
    num_copies = lambda x, y, z: x['variantCalls'][y]['numberOfCopies'][
        z
    ]['numberOfCopies']

    var_dict = add_columns_to_dict(columns)
    pb_idx = index_participant(variant, proband)
    var_dict["Chr"] = variant["coordinates"]["chromosome"]
    var_dict["Pos"] = variant["coordinates"]["start"]
    var_dict["End"] = variant["coordinates"]["end"]
    var_dict["Length"] = abs(var_dict["End"] - var_dict["Pos"])
    var_dict["Type"] = "STR"
    var_dict["Priority"] = "TIER1_STR"
    var_dict["Repeat"] = variant[
        "shortTandemRepeatReferenceData"
    ]["repeatedSequence"]
    var_dict["STR1"] = num_copies(variant, pb_idx, 0)
    var_dict["STR2"] = num_copies(variant, pb_idx, 1)
    var_dict["Gene"] = get_gene_symbol(variant)
    var_dict["AF Max"] = get_af_max(variant)
    return var_dict


def get_snv_info(variant, pb, ev_idx, columns, mother, father, pb_sex):
    '''
    Each variant that will be added to the excel workbook, needs to be
    added to the dataframe via a dictionary of values for each column
    heading in the workbook

    This function creates a variant dict for specific SNV to be added to
    the excel workbook, where the keys are the columns in the workbook and
    the values are the values for this variant.

    Inputs:
        variant: (dict) dict extracted from JSON describing single variant
        pb (str): participant ID of proband
        ev_index (int): index of reportEvents list for the event
        columns (list): list of columns to make into keys for variant dict
        mother (str): participant ID of mother
        father (str): participant ID of father
        pb_sex (str): sex of proband
    Outputs:
        var_dict: (dict) dict of variant information extracted from JSON
        will be added to a list of dicts for conversion into dataframe.
    '''
    var_dict = add_columns_to_dict(columns)
    mother_idx = index_participant(variant, mother)
    father_idx = index_participant(variant, father)
    pb_idx = index_participant(variant, pb)
    var_dict["Chr"] = variant["variantCoordinates"]["chromosome"]
    var_dict["Pos"] = variant["variantCoordinates"]["position"]
    var_dict["Ref"] = variant["variantCoordinates"]["reference"]
    var_dict["Alt"] = variant["variantCoordinates"]["alternate"]
    var_dict["Type"] = "SNV"
    var_dict["Priority"] = convert_tier(
        variant["reportEvents"][ev_idx]["tier"], "SNV"
    )
    var_dict["Zygosity"] = get_zygosity(
        zygosity=variant["variantCalls"][pb_idx]["zygosity"],
        p_sex=pb_sex,
        chrom=var_dict["Chr"]
    )
    
    var_dict["Depth"] = variant["variantCalls"][pb_idx]['depthAlternate']
    var_dict["Gene"] = get_gene_symbol(variant)
    var_dict['AF Max'] = get_af_max(variant)
    var_dict["Penetrance filter"] = variant["reportEvents"][ev_idx][
        "penetrance"
    ]
    var_dict["Panel MOI"] = convert_moi(
        variant["reportEvents"][ev_idx]["modeOfInheritance"]
    )
    var_dict["Inheritance"] = (
        get_inheritance(
            variant, mother_idx, father_idx, pb_sex
        )
    )
    return var_dict

def get_zygosity(zygosity, p_sex, chrom):
    '''
    Get the zygosity of the variant, and if the variant is heterozygous, 
    on the X chromosome and the proband is male then set the
    zygosity to hemizygous.

    Inputs:
        zygosity: (dict) dict extracted from JSON describing single variant.
        p_sex: (str) sex of the proband.
        chrom: (str) chromosome of the variant.

    Outputs:
        zygosity: (str) zygosity of the variant.

    '''

    if zygosity in ["heterozygous", "alternate_homozygous"] and \
        p_sex == "MALE" and chrom == "X":

        zygosity = "hemizygous"
    return zygosity



def get_cnv_info(variant, ev_index, columns):
    '''
    Each variant that will be added to the excel workbook, needs to be
    added to the dataframe via a dictionary of values for each column
    heading in the workbook

    This function creates a variant dict for specific CNV to be added to
    the excel workbook, where the keys are the columns in the workbook and
    the values are the values for this variant

    Inputs:
        variant: (dict) dict extracted from JSON describing single variant
        ev_index (int): index of reportEvents list for the event
        columns (list): list of columns to make into keys for variant dict
    Outputs:
        var_dict: (dict) dict of variant information extracted from JSON
        will be added to a list of dicts for conversion into dataframe.
    '''
    var_dict = add_columns_to_dict(columns)
    var_dict["Chr"] = variant["coordinates"]["chromosome"]
    var_dict["Pos"] = variant["coordinates"]["start"]
    var_dict["End"] = variant["coordinates"]["end"]
    var_dict["Length"] = abs(var_dict["End"] - var_dict["Pos"])
    var_dict["Type"] = "CNV"
    var_dict["Priority"] = convert_tier(
        variant["reportEvents"][ev_index]["tier"], "CNV"
    )
    var_dict["Copy Number"] = variant["variantCalls"][
        0
        ]['numberOfCopies'][0]['numberOfCopies']
    var_dict["Gene"] = get_gene_symbol(variant)
    var_dict['AF Max'] = get_af_max(variant)
    return var_dict


def get_top_3_ranked(df):
    '''
    Filter a df to return a df of the top 3 ranked Exomiser variants; this
    function uses a podium format so that equal ranks can be reported back.
    It will return all variants at each rank; so all the first ranked, all
    the second ranked and all the third ranked variants.
    Inputs
        df (pd.Dataframe): variant dataframe with a column containing
        strings of the exomiser ranks in the format "Exomiser Rank #"
    Outputs:
        df (pd.Dataframe): filtered variant dataframe containing only
        variants in the top three ranks
    '''
    # First change "Exomiser Rank #" string to int
    df['priority_as_int'] = df['Priority'].map(
        lambda x: int(x.split(' ')[-1])
    )
    # Get unique ranks and sort, selecting the top three ranks
    unique_ranks = df['priority_as_int'].unique()
    top_3_ranks = sorted(unique_ranks)[:3]

    # filter the df to include only values in top three ranks
    df = df[df['priority_as_int'].isin(top_3_ranks)]
    df.drop(['priority_as_int'], axis=1, inplace=True)
    return df


def look_up_id_in_refseq_mane_conversion_file(conversion, query_id, id_type):
    '''
    Search contents of a conversion file for a given ID. If a match is found,
    output the matched ID, else output None
        Inputs:
            query_id (str): ID to query (an ensembl ID)
            conversion (list): list of lines from a transcript
            nomenclature conversion file. Either the MANE file, which has only
            MANE information, or the RefSeq file which has all RefSeq
            transcripts
            id_type (str): string to match on e.g. NM to get the MANE
            transcript or ENSP to get the corresponding protein ID
        Outputs:
            matched_id (str): Matched ID or None if no match found.
    '''
    matched_id = None
    matches = []
    for line in conversion:
        if query_id in line:
            matches = [x for x in line.split() if x.startswith(id_type)]
            break

    if matches != []:
        if len(matches) == 1:
            matched_id = matches[0]
        else:
            print(
                f"Multiple matches for {query_id} found {', '.join(matches)}."
                "\nUnable to assign a match to this transcript."
            )

    return matched_id


def get_hgvs_exomiser(variant, mane, refseq_tsv):
    '''
    Exomiser variants have HGVS nomenclature for p dot and c dot provided
    in one field in the JSON in the following format:
    gene_symbol:ensembl_transcript_id:c_dot:p_dot
    This function extracts the cdot (with MANE refseq equivalent to
    ensembl transcript ID if found) and pdot (with ensembl protein ID)
    Exomiser variants only have one transcript in the JSON.
    Inputs:
        variant: (dict) dict describing single variant from JSON
        mane: MANE file for transcripts
        refseq: refseq file for transcripts
    Outputs:
        hgvs_c: (str) HGVS c dot (transcript) nomenclature for the variant.
        annotated against refseq transcript ID if one exists, and ensembl
        transcript ID if there is no matched equivalent
        hgvs_p: (str) HGVS p dot (protein) nomenclature for the variant.
        this is annotated against the ensembl protein ID.
    '''
    hgvs_c = None
    hgvs_p = None
    hgvs_source = variant['variantAttributes'][
        'additionalTextualVariantAnnotations'
        ]['hgvs']
    # Try converting Ensembl transcript to get RefSeq MANE
    refseq = look_up_id_in_refseq_mane_conversion_file(
        mane, hgvs_source.split(':')[1], "NM"
    )

    # If no MANE, return Ensembl transcript nomenclature
    if refseq is not None:
        hgvs_c = refseq + ":" + hgvs_source.split(':')[2]
    else:
        hgvs_c = hgvs_source.split(':')[1] + ':' + hgvs_source.split(':')[2]

    # get equivalent ENSP to transcript and construct p dot equivalent.
    ensp = look_up_id_in_refseq_mane_conversion_file(
        refseq_tsv, hgvs_source.split(':')[1].split('.')[0], "ENSP"
    )
    if ensp is not None:
        hgvs_p = ensp + ':' + hgvs_source.split(':')[3]

    return hgvs_c, hgvs_p


def get_hgvs_gel(variant, mane, refseq_tsv):
    '''
    GEL variants store HGVS p dot and c dot nomenclature separately.
    This function extracts the cdot (with MANE refseq equivalent to
    ensembl transcript ID if found) and pdot (with ensembl protein ID)
    Inputs:
        variant: (dict) dict describing single variant from JSON
        mane: MANE file for transcripts
        refseq: refseq file for transcripts
    Outputs:
        hgvs_c: (str) HGVS c dot (transcript) nomenclature for the variant.
        annotated against refseq transcript ID if one exists, and ensembl
        transcript ID if there is no matched equivalent
        hgvs_p: (str) HGVS p dot (protein) nomenclature for the variant.
        this is annotated against the ensembl protein ID.
    '''
    hgvs_p = None
    hgvs_c = None
    ref_list = []
    enst_list = []
    cdnas = variant['variantAttributes']['cdnaChanges']
    protein_changes = variant['variantAttributes']['proteinChanges']

    # Check for a MANE match
    for cdna in cdnas:
        refseq = look_up_id_in_refseq_mane_conversion_file(
            mane, cdna.split('(')[0], "NM"
        )

        if refseq is not None:
            ref_list.append(refseq + cdna.split(')')[1])
            enst_list.append(cdna.split('(')[0])

    # If no MANE match is found, return all transcripts
    if len(set(ref_list)) == 0:
        hgvs_c_list = []
        ensp_list = []
        for cdna in cdnas:
            hgvs_c_list.append(re.sub("\(.*?\)", "", cdna))
        for protein in protein_changes:
            ensp_list.append(protein)
        hgvs_c = ', '.join(hgvs_c_list)
        hgvs_p = ', '.join(ensp_list)

    else:
        hgvs_c = ', '.join(list(set(ref_list)))
        hgvs_p_list = []
        for enst in enst_list:
            ensp = look_up_id_in_refseq_mane_conversion_file(
                refseq_tsv, enst.split(".")[0], "ENSP"
            )
            if protein_changes != []:
                for protein in protein_changes:
                    if ensp in protein:
                        hgvs_p_list.append(protein)
        hgvs_p = ', '.join(hgvs_p_list)

    return hgvs_c, hgvs_p

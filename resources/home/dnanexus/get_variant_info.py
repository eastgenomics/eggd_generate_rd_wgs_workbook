import re


class VariantInfo():
    '''
    Functions to get variant data to be added to the excel spreadsheet
    '''
    @staticmethod
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

    @staticmethod
    def get_gene_symbol(variant):
        '''
        Get gene symbol from variant record
        Inputs:
            variant (dict): record for that specific variant from the JSON
        Outputs:
            gene_symbol (str): the gene symbol for that variant
        '''
        gene_list = []
        for entry in variant['reportEvents'][0]['genomicEntities']:
            if entry['type'] == 'gene':
                gene_list.append(entry['geneSymbol'])
        uniq_genes = list(set(gene_list))
        if len(uniq_genes) == 1:
            gene_symbol = uniq_genes[0]
        else:
            gene_symbol = str(uniq_genes).strip('[').strip(']')
        return gene_symbol

    @staticmethod
    def convert_tier(tier, var_type):
        '''
        Convert tier from GEL tiering variant to include variant type.
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
        elif tier == "TIERA" and var_type == "CNV":
            tier = "TIER1_CNV"
        elif tier == "TIER1" and var_type == "STR":
            tier = "TIER1_STR"
        return tier

    @staticmethod
    def get_af_max(variant):
        '''
        Get AF for population with highest allele frequency in the JSON
        Inputs:
            variant (dict): dict describing a single variant from JSON
        Outputs:
            highest_af: frequency of the variant in the population with
            the highest allele frequency of the variant, or 0 if the variant
            is not seen in any populations in the JSON
        '''
        highest_af = 0
        if variant['variantAttributes']['alleleFrequencies'] is not None:
            for af in variant['variantAttributes']['alleleFrequencies']:
                if af['alternateFrequency'] > highest_af:
                    highest_af = af['alternateFrequency']

        return highest_af

    @staticmethod
    def get_inheritance(variant, m_idx, f_idx, p_sex):
        '''
        Work out inheritance of variant based on zygosity of parents
        Inputs:
            variant: (dict): dict extracted from JSON describing single variant
            m_idx: (int): position within list of variant calls that
            belongs to the mother, or None if no mother in JSON
            f_idx: (int): position within list of variant calls that
            belongs to the father, or None if no father in JSON
            p_sex (str): proband sex
        Outputs:
            inheritance (str): inferred inheritance of the variant, or None.
        '''
        zygosity = lambda x, y: x['variantCalls'][y]['zygosity']

        inheritance = None
        maternal = False
        paternal = False

        inheritance_types = ['alternate_homozygous', 'heterozygous']

        if m_idx is not None:
            if zygosity(variant, m_idx) in inheritance_types:
                maternal = True

        if f_idx is not None:
            if zygosity(variant, f_idx) in inheritance_types:
                if (p_sex == 'MALE' and
                variant['variantCoordinates']['chromosome'] == 'X'):
                    paternal = False
                else:
                    paternal = True

        if maternal is True and paternal is False:
            inheritance = "maternal"

        elif maternal is False and paternal is True:
            inheritance = "paternal"

        elif maternal is True and paternal is True:
            inheritance = "both"

        return inheritance

    @staticmethod
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

    @staticmethod
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
                    f"Unable to find proband ID {participant_id} in variant"
                    f"{variant['variantCalls']}"
                )

        return index

    @staticmethod
    def get_str_info(variant, proband, columns):
        '''
        Fill in variant dict for specific STR variant
        Inputs:
            variant (dict): dict extracted from JSON describing single variant
            proband (str): GEL ID for the proband
            columns (list): list of columns to make into keys for variant dict
        Outputs:
            var_dict: (dict) dict of variant information extracted from JSON
            will be added to a list of dicts for conversion into dataframe.
        '''
        num_copies = lambda x, y, z: x['variantCalls'][y]['numberOfCopies'][
            z
            ]['numberOfCopies']

        var_dict = VariantInfo.add_columns_to_dict(columns)
        pb_idx = VariantInfo.index_participant(variant, proband)
        var_dict["Chr"] = variant["coordinates"]["chromosome"]
        var_dict["Pos"] = variant["coordinates"]["start"]
        var_dict["End"] = variant["coordinates"]["end"]
        var_dict["Length"] = abs(var_dict["End"] - var_dict["Pos"])
        var_dict["Type"] = "STR"
        var_dict["Priority"] = "STR"
        var_dict["Repeat"] = variant[
            "shortTandemRepeatReferenceData"
        ]["repeatedSequence"]
        var_dict["STR1"] = num_copies(variant, pb_idx, 0)
        var_dict["STR2"] = num_copies(variant, pb_idx, 1)
        var_dict["Gene"] = VariantInfo.get_gene_symbol(variant)
        return var_dict

    @staticmethod
    def get_snv_info(variant, pb, ev_idx, columns, mother, father, pb_sex):
        '''
        Fill in variant dict for specific SNV
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
        var_dict = VariantInfo.add_columns_to_dict(columns)
        m_idx = VariantInfo.index_participant(variant, mother)
        f_idx = VariantInfo.index_participant(variant, father)
        pb_idx = VariantInfo.index_participant(variant, pb)
        var_dict["Chr"] = variant["variantCoordinates"]["chromosome"]
        var_dict["Pos"] = variant["variantCoordinates"]["position"]
        var_dict["Ref"] = variant["variantCoordinates"]["reference"]
        var_dict["Alt"] = variant["variantCoordinates"]["alternate"]
        var_dict["Type"] = "SNV"
        var_dict["Priority"] = VariantInfo.convert_tier(
            variant["reportEvents"][ev_idx]["tier"], "SNV"
        )
        var_dict["Zygosity"] = variant["variantCalls"][pb_idx]["zygosity"]
        var_dict["Depth"] = variant["variantCalls"][pb_idx]['depthAlternate']
        var_dict["Gene"] = VariantInfo.get_gene_symbol(variant)
        var_dict['AF Max'] = VariantInfo.get_af_max(variant)
        var_dict["Penetrance filter"] = variant["reportEvents"][ev_idx][
            "penetrance"
        ]
        print(
            variant["reportEvents"][ev_idx]["modeOfInheritance"]
        )
        var_dict["Inheritance mode"] = VariantInfo.convert_moi(
            variant["reportEvents"][ev_idx]["modeOfInheritance"]
        )
        var_dict["Inheritance"] = (
            VariantInfo.get_inheritance(variant, m_idx, f_idx, pb_sex)
        )
        return var_dict

    @staticmethod
    def get_cnv_info(variant, ev_index, columns):
        '''
        Extract CNV info from JSON and create variant dict of required data for
        workbook
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            ev_index (int): index of reportEvents list for the event
            columns (list): list of columns to make into keys for variant dict
        Outputs:
            var_dict: (dict) dict of variant information extracted from JSON
            will be added to a list of dicts for conversion into dataframe.
        '''
        var_dict = VariantInfo.add_columns_to_dict(columns)
        var_dict["Chr"] = variant["coordinates"]["chromosome"]
        var_dict["Pos"] = variant["coordinates"]["start"]
        var_dict["End"] = variant["coordinates"]["end"]
        var_dict["Length"] = abs(var_dict["End"] - var_dict["Pos"])
        var_dict["Type"] = "CNV"
        var_dict["Priority"] = VariantInfo.convert_tier(
            variant["reportEvents"][ev_index]["tier"], "CNV"
        )
        var_dict["Copy Number"] = variant["variantCalls"][
            0
            ]['numberOfCopies'][0]['numberOfCopies']
        var_dict["Gene"] = VariantInfo.get_gene_symbol(variant)
        var_dict['AF Max'] = VariantInfo.get_af_max(variant)
        return var_dict

    @staticmethod
    def get_top_3_ranked(ranked):
        '''
        Get top 3 ranked Exomiser variants; this function uses a podium format
        so that equal ranks can be reported back.
        Uses gold, silver and bronze to refer to the top, second and third
        ranked times
        Inputs
            ranked (list): list of Exomiser variants
        Outputs:
            to_report (list): top three Exomiser variants to report back
        '''
        # TODO: Reconfigure based on analyst feedback
        # Debate between “olympic-style” podium and “boxing-style” podium.
        # i.e. should this be 1 2 2 3 == 1 2 2 or 1 2 2 3
        gold = []
        silver = []
        bronze = []

        rank = lambda x: x['reportEvents']['vendorSpecificScores']['rank']

        ordered_list = sorted(ranked, key=rank)

        for snv in ordered_list:
            if not gold:
                gold = [snv]
                continue
            else:
                if rank(snv) == rank(gold[0]):
                    gold.append(snv)
                    continue

            if len(gold) >= 3:
                break

            if not silver:
                silver = [snv]
                continue
            else:
                if rank(snv) == rank(silver[0]):
                    silver.append(snv)
                    continue

            if len(gold) + len(silver) >= 3:
                break

            if not bronze:
                bronze = [snv]
                continue
            else:
                if rank(snv) == rank(bronze[0]):
                    bronze.append(snv)
                    continue
                else:
                    break

        to_report = gold + silver + bronze

        return to_report


class VariantNomenclature():
    '''
    Functions for manipulating variant nomenclature.
    '''
    @staticmethod
    def convert_ensembl_to_refseq(mane, ensembl):
        '''
        Search MANE file for query ensembl transcript ID. If match found,
        output the equivalent refseq ID, else output None
        Inputs:
            ensembl (str): Ensembl transcript ID from JSON
            mane (list): list of lines from MANE file.
        Outputs:
            refseq (str): Matched MANE refseq ID to Ensembl transcript ID, or
            None if no match found.
        '''
        refseq = None

        for line in mane:
            if ensembl in line:
                refseq = [x for x in line.split() if x.startswith('NM')]
                refseq = refseq[0]
                break
        return refseq

    @staticmethod
    def get_ensp(refseq, enst):
        '''
        from input ensembl transcript ID, get ensembl protein ID from the
        refseq_tsv
        Inputs:
            enst (str): Ensembl transcript ID from JSON
            refseq (list): list of lines from refseq tsv file.
        Outputs:
            ensp (str): Ensembl protein ID equivalent to the input transcript
            ID
        '''
        ensp = None
        for line in refseq:
            if enst in line:
                ensp = [x for x in line.split() if x.startswith('ENSP')]
                ensp = ensp[0]
                break
        return ensp

    @staticmethod
    def get_hgvs_exomiser(variant, mane, refseq_tsv):
        '''
        Exomiser variants have HGVS nomenclature for p dot and c dot provided
        in one field in the JSON in the following format:
        gene_symbol:ensembl_transcript_id:c_dot:p_dot
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
        hgvs_c = None
        hgvs_p = None
        hgvs_source = variant['variantAttributes'][
            'additionalTextualVariantAnnotations'
            ]['hgvs']
        refseq = VariantNomenclature.convert_ensembl_to_refseq(
            mane, hgvs_source.split(':')[1]
        )
        if refseq is not None:
            hgvs_c = refseq + ":" + hgvs_source.split(':')[2]
        ensp = VariantNomenclature.get_ensp(
            refseq_tsv, hgvs_source.split(':')[1].split('.')[0]
        )
        if ensp is not None:
            hgvs_p = ensp + ':' + hgvs_source.split(':')[3]
        return hgvs_c, hgvs_p

    @staticmethod
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

        for cdna in cdnas:
            refseq = VariantNomenclature.convert_ensembl_to_refseq(
                mane, cdna.split('(')[0]
            )

            if refseq is not None:
                ref_list.append(refseq + cdna.split(')')[1])
                enst_list.append(cdna.split('(')[0])

        if len(set(ref_list)) > 1:
            raise RuntimeError(
                f"Transcript {cdnas} matched more than one MANE transcript"
            )
        elif len(set(ref_list)) == 0:
            # if there is no MANE equivalent refseq, use the ensembl
            # transcript ID, remove gene ID in brackets.
            # currently we just use the first one in the list! is there a
            # better way to do this?
            # TODO: this will considered and corrected with feedback
            # TODO: provide all of them?
            hgvs_c = re.sub("\(.*?\)", "", cdnas[0])
            ensp = VariantNomenclature.get_ensp(
                refseq_tsv, cdnas[0].split('(')[0]
            )
            for protein in protein_changes:
                if ensp in protein:
                    hgvs_p = protein

        else:
            hgvs_c = list(set(ref_list))[0]
            ensp = VariantNomenclature.get_ensp(refseq_tsv, enst_list[0])
            for protein in protein_changes:
                if ensp in protein:
                    hgvs_p = protein

        return hgvs_c, hgvs_p

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
        '''
        variant_dict = {}
        for col in column_list:
            variant_dict[col] = ""

        return variant_dict
    
    @staticmethod
    def get_gene_symbol(variant):
        '''
        Get gene symbol from variant record
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
        Convert tier to add SNV as CNVs are also Tier 1
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
        elif tier == "TIER2" and var_type == "STR":
            tier = "TIER2_STR"
            # TODO fix this; there are no TIER2 STRs!
            # TODO write in readme how var filtering works!
        return tier
    
    @staticmethod
    def get_af_max(variant):
        '''
        Get AF for population with highest allele frequency in the JSON
        Inputs:
            variant (dict): dict describe a single variant from JSON
        Outputs:
            highest_af (int): frequency of the variant in the population with
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
    def get_inheritance(variant, proband_index):
        '''
        Work out inheritance of variant based on zygosity of parents
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            proband_index: (int) position within list of variant calls that
            belongs to the proband
        Outputs:
            inheritance (str): inferred inheritance of the variant, or None.
        '''
        # TODO entirely reconfigure this based on what GEL said!
        inheritance = None
        mother_index = None
        father_index = None
        inheritance_types = ['alternate_homozygous', 'heterozygous']
        calls = variant['variantCalls']

        # if self.mother and self.father:
        #     # find index of variant call list for mother and father            
        #     for call in variant['variantCalls']:
        #         print(call)
        #         if call['participantId'] == self.mother:
        #             mother_index = variant['variantCalls'].index(call)
        #         elif call['participantId'] == self.father:
        #             father_index = variant['variantCalls'].index(call)
        #     if mother_index and father_index:
        #         if calls[proband_index]['zygosity'] in inheritance_types:

        #             if (calls[
        #                     mother_index
        #                     ]['zygosity'] == 'reference_homozygous'
        #                 and calls[
        #                     father_index
        #                     ]['zygosity'] in inheritance_types
        #                 ):
        #                 inheritance = "paternal"

        #             elif (calls[
        #                     mother_index
        #                     ]['zygosity'] in inheritance_types
        #                 and calls[
        #                     father_index
        #                     ]['zygosity'] == 'reference_homozygous'
        #                 ):
        #                 inheritance = "maternal"
        #     else:
        #         confusing_call = calls[proband_index]['zygosity']
        #         print(
        #             f"proband call {confusing_call} not recognised"
        #         )
        #         # can we reliable say this? it could have been inherited from the father
        # # we just don't have his data?
        # elif self.mother and not self.father:
        #     for call in variant['variantCalls']:
        #         if call['participantId'] == self.mother:
        #             mother_index = variant['variantCalls'].index(call)
        #     if calls[mother_index]['zygosity'] in inheritance_types:
        #         inheritance = "maternal"
        # # same as above?
        # elif self.father and not self.mother:
        #     for call in variant['variantCalls']:
        #         if call['participantId'] == self.father:
        #             father_index = variant['variantCalls'].index(call)

        #     if calls[father_index]['zygosity'] in inheritance_types:
        #         inheritance = "paternal"

        return inheritance

    @staticmethod
    def get_str_info(variant, proband_index, columns):
        '''
        Fill in variant dict for specific STR variant
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            proband_index: (int) position within list of variant calls that
            belongs to the proband
        Outputs:
            var_dict: (dict) dict of variant information extracted from JSON
            will be added to a list of dicts for conversion into dataframe.
        '''
        var_dict = VariantInfo.add_columns_to_dict(columns)
        var_dict["Chr"] = variant["coordinates"]["chromosome"]
        var_dict["Pos"] = variant["coordinates"]["start"]
        var_dict["End"] = variant["coordinates"]["end"]
        var_dict["Length"] = abs(var_dict["End"] - var_dict["Pos"])
        var_dict["Type"] = "STR"
        var_dict["Priority"] = "STR"
        var_dict["Repeat"] = variant[
            "shortTandemRepeatReferenceData"
        ]["repeatedSequence"]
        var_dict["STR1"] = variant['variantCalls'][proband_index][
            'numberOfCopies'
        ][0]['numberOfCopies']
        var_dict["STR2"] = variant['variantCalls'][proband_index][
            'numberOfCopies'
        ][1]['numberOfCopies']
        var_dict["Gene"] = VariantInfo.get_gene_symbol(variant)
        return var_dict

    @staticmethod
    def get_snv_info(variant, pb_index, ev_index, columns):
        '''
        Fill in variant dict for specific SNV
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            pb_index (int): index of variantCalls list for the proband
            ev_index (int): index of reportEvents list for the event
        Outputs:
            var_dict: (dict) dict of variant information extracted from JSON
            will be added to a list of dicts for conversion into dataframe.
        '''
        var_dict = VariantInfo.add_columns_to_dict(columns)
        var_dict["Chr"] = variant["variantCoordinates"]["chromosome"]
        var_dict["Pos"] = variant["variantCoordinates"]["position"]
        var_dict["Ref"] = variant["variantCoordinates"]["reference"]
        var_dict["Alt"] = variant["variantCoordinates"]["alternate"]
        var_dict["Type"] = "SNV"
        var_dict["Priority"] = VariantInfo.convert_tier(
            variant["reportEvents"][ev_index]["tier"], "SNV"
        )
        var_dict["Zygosity"] = variant["variantCalls"][pb_index]["zygosity"]
        var_dict["Depth"] = variant["variantCalls"][pb_index]['depthAlternate']
        var_dict["Gene"] = VariantInfo.get_gene_symbol(variant)
        var_dict['AF Max'] = VariantInfo.get_af_max(variant)
        var_dict["Penetrance filter"] = variant["reportEvents"][ev_index]["penetrance"]
        var_dict["Inheritance mode"] = variant["reportEvents"][ev_index][
            "modeOfInheritance"
        ]
        var_dict["Segregation pattern"] = variant["reportEvents"][ev_index][
            "segregationPattern"
        ]
        var_dict["Inheritance"] = VariantInfo.get_inheritance(variant, pb_index)
        return var_dict

    @staticmethod
    def get_cnv_info(variant, ev_index, columns):
        '''
        Extract CNV info from JSON and create variant dict of required data for
        workbook
        Inputs:
            variant: (dict) dict extracted from JSON describing single variant
            ev_index (int): index of reportEvents list for the event
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
                refseq = line.split(',')[3].strip("\"")
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

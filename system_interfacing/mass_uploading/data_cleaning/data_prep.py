import pandas as pd
import json

from ..helpers.utils.utils import get_dates


def load_json(file_path):
    with open(file_path) as json_data:
        mtg_dict = json.load(json_data)
    return mtg_dict


def filter_keys(list_of_dictionaries,needed_keys_list):
    dict1 = []
    for d in list_of_dictionaries:
        filtered_d = dict((k, d[k]) for k in needed_keys_list if k in d)
        dict1.append(filtered_d)
    return dict1


def deduplicate_filtered_keys(filtered_list_of_dictionaries):
    dict2 = []
    for d in filtered_list_of_dictionaries:
        if d not in dict2:
            dict2.append(d)
    return dict2


def get_cpm_dict(upload_posting_folder, upload_posting_folder_test, target_file):
    df = pd.read_excel(f"{upload_posting_folder}{target_file}")
    df = df[~df['BRANCH'].str.contains('unknown')]
    df = df[['Name', 'DOC_GROUP', 'DOC_TYPE', 'PROTOCOL', 'BRANCH', 'MEETING']]
    dl_to_protnames = {'name1': 'name2'}
    df['ProtocolName'] = df['PROTOCOL'].replace(dl_to_protnames, regex=True)
    df_cmt = pd.read_excel(upload_posting_folder_test + 'prot_cmt_type.xlsx')
    df_cmt.columns = ['ProtocolName', 'cmt_type']
    df_cpm = pd.read_excel(upload_posting_folder_test + 'cpm_branch.xlsx')
    cpm_dict = df_cpm.to_dict('records')
    return df, df_cmt, cpm_dict


def get_mtg_dict_narrowed(datapath, df, df_cmt):
    df_branch_adjusted = pd.read_excel(f"{datapath}prot_branch_adjusted.xlsx")
    df_branch_adjusted.columns = ['ProtocolName', 'cpm_branch']
    df1 = df.merge(df_cmt, on='ProtocolName', how='left')
    df1['mtg_date'] = df1['MEETING'].apply(lambda x: get_dates(x))
    df1 = df1.merge(df_branch_adjusted, on='ProtocolName', how='left')
    docs_final_dict = df1.to_dict('records')

    # JSON of meeting data to associate with documents to-be-migrated
    mtg_dict = load_json(f"{datapath}jsons/meeting_json_dump.txt")
    output_dict1 = [x for x in mtg_dict["Data"]]
    mtg_dict_narrowed = deduplicate_filtered_keys(
        filter_keys(
            output_dict1, ['MeetingID', 'MeetingName']
        )
    )
    return mtg_dict_narrowed, docs_final_dict


def get_prepped_data(target_file, datapath, upload_posting_folder_test, upload_posting_folder):
    df, df_cmt, cpm_dict = get_cpm_dict(upload_posting_folder, upload_posting_folder_test, target_file)
    mtg_dict_narrowed, docs_final_dict = get_mtg_dict_narrowed(datapath, df, df_cmt)
    return mtg_dict_narrowed, docs_final_dict, cpm_dict, df



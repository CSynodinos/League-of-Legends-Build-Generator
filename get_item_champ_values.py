import requests
import pandas as pd
import numpy as np
import csv
import os

def data_version():
    ddragon = "https://ddragon.leagueoflegends.com/realms/euw.json"
    euw_json = requests.get(ddragon).json()
    return euw_json['n']['champion']


def build_champ_data_url():
    return "http://ddragon.leagueoflegends.com/cdn/" + data_version() + "/data/en_GB/champion.json"

def build_item_data_url():
    return "http://ddragon.leagueoflegends.com/cdn/" + data_version() + "/data/en_GB/item.json"

def get_champ_json():
    data_url = build_champ_data_url()
    data_json = requests.get(data_url).json()
    champ_list = data_json['data'].keys()
    return data_json, champ_list

def get_item_json():
    data_url = build_item_data_url()
    data_json = requests.get(data_url).json()
    item_list = data_json['data'].keys()
    return data_json, item_list

def champ_headings():
    return [
        "Champion",
        "HP",
        "HPg",
        "MP",
        "MPg",
        "AD",
        "ADg",
        "AS",
        "ASg",
        "AR",
        "ARg",
        "MR",
        "MRg",
        "MS"]

def item_headings():
    return [
        "Item",
        "AP",
        "AD",
        "AS",
        "HP",
        "MP",
        "AR",
        "MR",
        "Critical",
        "Life Steal",
        "Haste",
        "MS"]

def champions():
    data_json, champ_list = get_champ_json()
    file_name = 'Champion_Stats.csv'
    with open(file_name, 'w', newline='', encoding='utf8') as csv_file:
        writer = csv.writer(csv_file, delimiter=',')
        writer.writerow(champ_headings())
        for champ in champ_list:
            name = data_json['data'][champ]['name']
            hp = data_json['data'][champ]['stats']['hp']
            hpperlevel = data_json['data'][champ]['stats']['hpperlevel']
            mp = data_json['data'][champ]['stats']['mp']
            mpperlevel = data_json['data'][champ]['stats']['mpperlevel']
            ad = data_json['data'][champ]['stats']['attackdamage']
            adg = data_json['data'][champ]['stats']['attackdamageperlevel']
            aspeed = data_json['data'][champ]['stats']['attackspeed']
            asg = data_json['data'][champ]['stats']['attackspeedperlevel']
            ar = data_json['data'][champ]['stats']['armor']
            arg = data_json['data'][champ]['stats']['armorperlevel']
            mr = data_json['data'][champ]['stats']['spellblock']
            mrg = data_json['data'][champ]['stats']['spellblockperlevel']
            ms= data_json['data'][champ]['stats']['movespeed']

            writer.writerow([name, hp, hpperlevel, mp, mpperlevel, ad, adg, aspeed, asg, ar, arg, mr, mrg, ms])
        
    read_file = pd.read_csv ('Champion_Stats.csv')
    GFG = pd.ExcelWriter('Champion_Stats.xlsx')
    read_file.to_excel(GFG, index = False)
    GFG.save()
    os.remove('Champion_Stats.csv')



######### need to add the numbers for each item.
def items():
    data_json, item_list = get_item_json()
    file_name = 'Item_Stats.csv'
    
    with open(file_name, 'w', newline='', encoding='utf8') as csv_file:
        writer = csv.writer(csv_file, delimiter=',')
        writer.writerow(item_headings())
        for item in item_list:
            item = data_json['data'][item]['name']
            writer.writerow([item])


def main():
    champions()
    items()

if __name__ == "__main__":
    main()
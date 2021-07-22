import requests
import pandas as pd
import csv
import os

def data_version():
    """Finds game version using the Riot Games API."""
    
    server_json = "https://ddragon.leagueoflegends.com/realms/euw.json"
    euw_json = requests.get(server_json).json()
    
    return euw_json['n']['champion']

class api():
    """Gets champion and item data urls using the Riot Games API."""
    
    def __init__(self, dragon, champurl, itemurl):
        self.dragon = dragon
        self.champurl = champurl
        self.itemurl = itemurl
        
    def build_champ_data_url(self):
        return self.dragon + data_version() + self.champurl

    def build_item_data_url(self):
        return self.dragon + data_version() + self.itemurl

d = api("http://ddragon.leagueoflegends.com/cdn/", "/data/en_GB/champion.json", "/data/en_GB/item.json") 

def get_champ_json():
    """Gets the champion json from the current game version url."""
    
    data_url = d.build_champ_data_url()
    data_json = requests.get(data_url).json()
    champ_list = data_json['data'].keys()
    
    return data_json, champ_list

def get_item_json():
    """Gets the item json from the current game version url."""
    
    data_url = d.build_item_data_url()
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
    """Writes champion data into csv and converts into xlsx from json"""
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

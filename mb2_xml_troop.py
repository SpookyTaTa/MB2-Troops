import xml.etree.ElementTree as ET
import openpyxl
from openpyxl import Workbook, load_workbook
import pandas as pd
from pandas import ExcelWriter
import csv


def get_armor(xml_file, num_of_ratings, armor_type_1, armor_type_2, armor_type_3):
	tree = ET.parse(xml_file)
	root = tree.getroot()

	ar_2D_list = []

	for ar in root.findall('Item'):
		id_ar = ar.get('id')
		name = ar.get('name').partition("}")[2]
		culture = ar.get('culture').partition(".")[2]
		if culture == "neutral_culture":
			culture = "neutral"
		weight = ar.get('weight')
		
		if num_of_ratings == 3:
			ar_rating_1 = ar.find('ItemComponent').find('Armor').get(armor_type_1)
			ar_rating_2 = ar.find('ItemComponent').find('Armor').get(armor_type_2)
			ar_rating_3 = ar.find('ItemComponent').find('Armor').get(armor_type_3)
			entry = [id_ar, name, culture, ar_rating_1, ar_rating_2, ar_rating_3, weight]
		elif num_of_ratings == 2:
			ar_rating_1 = ar.find('ItemComponent').find('Armor').get(armor_type_1)
			ar_rating_2 = ar.find('ItemComponent').find('Armor').get(armor_type_2)
			entry = [id_ar, name, culture, ar_rating_1, ar_rating_2, weight]
		elif num_of_ratings == 1:
			ar_rating_1 = ar.find('ItemComponent').find('Armor').get(armor_type_1)
			entry = [id_ar, name, culture, ar_rating_1, weight]
		
		entry = ["0" if i == None else i for i in entry]
		print(entry)
		ar_2D_list.append(entry)

	file = open('mb2.csv', 'w', newline ='')

	with file:
		write = csv.writer(file)
		write.writerows(ar_2D_list)

	df = pd.read_csv("mb2.csv")

	ws_name = xml_file.partition('_')[0]

	with ExcelWriter('MB2.xlsx', mode='a') as writer:
		at1 = armor_type_1.partition("_")
		at1 = at1[0].capitalize() + " " + at1[2].capitalize()
		if num_of_ratings >= 2:
			at2 = armor_type_2.partition("_")
			at2 = at2[0].capitalize() + " " + at2[2].capitalize()
			if num_of_ratings == 3:
				at3 = armor_type_3.partition("_")
				at3 = at3[0].capitalize() + " " + at3[2].capitalize()
				df.to_excel(writer, sheet_name=ws_name, index=False, header=["ID", "Name", "Culture", at1, at2, at3, "Weight"])
			elif num_of_ratings == 2:
				df.to_excel(writer, sheet_name=ws_name, index=False, header=["ID", "Name", "Culture", at1, at2,"Weight"])
		elif num_of_ratings == 1:
			df.to_excel(writer, sheet_name=ws_name, index=False, header=["ID", "Name", "Culture", at1, "Weight"])

# def get_npc(xml_file):
# 	tree = ET.parse(xml_file)
# 	root = tree.getroot()

# 	npc_2D_list = []

# 	for npc in root.findall('NPCCharacter'):
# 		id_npc = npc.get('id')
# 		culture = npc.get('culture').partition(".")[2]
# 		name = npc.get('name').partition("}")[2]
# 		troop_type = npc.get('default_group')
# 		occupation = npc.get('occupation')

# 		npc_2D_list = [id_npc, culture, name, troop_type, occupation]

# 		skills = npc.find('skills')
# 		for sk in skills.findall('skill')
# 			val = sk.get('value')
# 			npc_2D_list.append(val)

		

if __name__ == '__main__':
	get_armor('head_ar.xml', 1, 'head_armor', None, None)
	get_armor('shoulder_ar.xml', 2, 'body_armor', 'arm_armor', None)
	get_armor('body_ar.xml', 3, 'body_armor', 'arm_armor', 'leg_armor')
	get_armor('arm_ar.xml', 1, 'arm_armor', None, None)
	get_armor('leg_ar.xml', 1, 'leg_armor', None, None)

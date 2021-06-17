import xml.etree.ElementTree as ET
import openpyxl
from openpyxl import Workbook, load_workbook
import pandas as pd
from pandas import ExcelWriter
import csv
import numpy as np


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

	file = open('mb2.csv', 'w', newline ='', header=None)

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
				df.to_excel(writer, sheet_name=ws_name, index=False, header=["ID", "Name", "Culture", at1, at2, "Weight"])
		elif num_of_ratings == 1:
			df.to_excel(writer, sheet_name=ws_name, index=False, header=["ID", "Name", "Culture", at1, "Weight"])

def get_npc(xml_file):
	tree = ET.parse(xml_file)
	root = tree.getroot()

	npc_2D_list = []

	for npc in root.findall('NPCCharacter'):
		id_npc = npc.get('id')
		culture = npc.get('culture').partition(".")[2]
		name = npc.get('name').partition("}")[2]
		troop_type = npc.get('default_group')
		occupation = npc.get('occupation')

		entry = [id_npc, culture, name, troop_type, occupation]

		skills = npc.find('skills')
		skill_dict = {}
		# for sk in skills.findall('skill'):
		# 	sk_id = sk.get('id')
		# 	val = sk.get('value')
		# 	entry.append(val)
		print(name)
		for sk in skills.findall('skill'):
			sk_id = sk.get('id')
			print(sk_id)
			val = sk.get('value')
			print(val)
			skill_dict[sk_id] = val

		entry.append(skill_dict['OneHanded'])
		entry.append(skill_dict['TwoHanded'])
		entry.append(skill_dict['Polearm'])
		entry.append(skill_dict['Bow'])
		entry.append(skill_dict['Crossbow'])
		entry.append(skill_dict['Throwing'])
		entry.append(skill_dict['Riding'])
		entry.append(skill_dict['Athletics'])



		equipments = npc.find('Equipments')
		for eq_roster in equipments.findall('EquipmentRoster'):
			for eq in eq_roster:
				slot = eq.get('slot')
				id_eq = eq.get('id').partition(".")[2]
				entry.append(slot + "_" + id_eq)

		npc_2D_list.append(entry)


	df = pd.DataFrame(npc_2D_list)

	ws_name = xml_file.partition('.')[0]
	print(ws_name)

	with ExcelWriter('MB2.xlsx', mode='a') as writer:
		df.to_excel(writer, sheet_name=ws_name, index=False)


def convert_id_to_name(xml_file):
	tree = ET.parse(xml_file)
	root = tree.getroot()
	wpn_dict = {}

	for child in root:
		weapon_id = child.get('id')
		weapon_name = child.get('name').partition('}')[2]
		wpn_dict[weapon_id] = weapon_name

	print(wpn_dict)

	df = pd.read_excel('MB2.xlsx', sheet_name='npcs')
	for rowIndex, row in df.iterrows():
		for columnIndex, value in row.items():
			if type(value) == str:
				value_tag = value.partition('_')[0]
				value_id = value.partition('_')[2]
				if value_tag in ['Item0', 'Item1', 'Item2', 'Item3', 'Head', 'Cape', 'Body','Gloves', 'Leg']:
					df.at[rowIndex, columnIndex] = wpn_dict.get(value_id)

	with ExcelWriter('MB2.xlsx', mode='a') as writer:
		df.to_excel(writer, sheet_name='npcRevised', index=False)

def get_npc_armor_avg():
	# ACCESS EVERY SHEET IN THE XLSX FILE
	df_npc = pd.read_excel('MB2.xlsx', sheet_name='npcs')
	df_head = pd.read_excel('MB2.xlsx', sheet_name='head')
	df_shld = pd.read_excel('MB2.xlsx', sheet_name='shoulder')
	df_body = pd.read_excel('MB2.xlsx', sheet_name='body')
	df_arm = pd.read_excel('MB2.xlsx', sheet_name='arm')
	df_leg = pd.read_excel('MB2.xlsx', sheet_name='leg')
	troop_2D_list = []

	# FOR EVERY TROOP
	for rowIndex, row in df_npc.iterrows():
		head, shld, body, arm, leg = ([] for i in range(5))
		name = row[2]

		# FOR EVERY ITEM THAT A TROOP CAN HAVE
		for columnIndex, value in row.items():
			if type(value) == str:
				value_tag = value.partition('_')[0]
				value_id = value.partition('_')[2]

				# HEAD ARMOR
				if value_tag == 'Head':
					Head_head_armor = df_head.loc[df_head.ID == value_id, 'HeadArmor'].to_numpy()
					Head_head_armor = float(Head_head_armor)

					weight = df_head.loc[df_head.ID == value_id, 'Weight'].to_numpy()
					weight = float(weight)

					head.append([Head_head_armor, weight])
				# SHOULDER ARMOR
				elif value_tag == 'Cape':
					Shoulder_body_armor = df_shld.loc[df_shld.ID == value_id, 'Body Armor'].to_numpy()
					Shoulder_body_armor = float(Shoulder_body_armor)
					if Shoulder_body_armor is None:
						Shoulder_body_armor = 0

					Shoulder_arm_armor = df_shld.loc[df_shld.ID == value_id, 'Arm Armor'].to_numpy()
					Shoulder_arm_armor = float(Shoulder_arm_armor)
					if Shoulder_arm_armor is None:
						Shoulder_arm_armor = 0

					weight = df_shld.loc[df_shld.ID == value_id, 'Weight'].to_numpy()
					weight = float(weight)

					shld.append([Shoulder_body_armor, Shoulder_arm_armor, weight])
				# BODY ARMOR
				elif value_tag == 'Body':
					Body_body_armor = df_body.loc[df_body.ID == value_id, 'Body Armor'].to_numpy()
					Body_body_armor = float(Body_body_armor)
					if Body_body_armor is None:
						Body_body_armor = 0

					Body_arm_armor = df_body.loc[df_body.ID == value_id, 'Arm Armor'].to_numpy()
					Body_arm_armor = float(Body_arm_armor)
					if Body_arm_armor is None:
						Body_arm_armor = 0

					Body_leg_armor = df_body.loc[df_body.ID == value_id, 'Leg Armor'].to_numpy()
					Body_leg_armor = float(Body_leg_armor)
					if Body_leg_armor is None:
						Body_leg_armor = 0

					weight = df_body.loc[df_body.ID == value_id, 'Weight'].to_numpy()
					weight = float(weight)

					body.append([Body_body_armor, Body_arm_armor, Body_leg_armor, weight])
				# ARM ARMOR
				elif value_tag == 'Gloves':
					Arm_arm_armor = df_arm.loc[df_arm.ID == value_id, 'Arm Armor'].to_numpy()
					Arm_arm_armor = float(Arm_arm_armor)

					weight = df_arm.loc[df_arm.ID == value_id, 'Weight'].to_numpy()
					weight = float(weight)

					arm.append([Arm_arm_armor, weight])
				# LEG ARMOR
				elif value_tag == 'Leg':
					Leg_leg_armor = df_leg.loc[df_leg.ID == value_id, 'Leg Armor'].to_numpy()
					Leg_leg_armor = float(Leg_leg_armor)

					weight = df_leg.loc[df_leg.ID == value_id, 'Weight'].to_numpy()
					weight = float(weight)

					leg.append([Leg_leg_armor, weight])

		head_armor = body_armor = arm_armor = leg_armor = 0
		total_weight = 0

		# AVERAGE HEAD ARMOR STATS
		h = wt = 0
		for i in head:
			h += i[0]
			wt += i[1]
		try:
			head_armor += h / len(head)
			total_weight += wt / len(head)
		except ZeroDivisionError:
			pass

		# AVERAGE SHOULDER ARMOR STATS
		b = a = wt = 0
		for i in shld:
			b += i[0]
			a += i[1]
			wt += i[2]
		try:
			body_armor += b / len(shld)
			arm_armor += a / len(shld)
			total_weight += wt / len(shld)
		except ZeroDivisionError:
			pass

		# AVERAGE BODY ARMOR STATS
		b = a = l = wt = 0
		for i in body:
			b += i[0]
			a += i[1]
			l += i[2]
			wt += i[3]
		try:
			body_armor += b / len(body)
			arm_armor += a / len(body)
			leg_armor += l / len(body)
			total_weight += wt / len(body)
		except ZeroDivisionError:
			pass

		# AVERAGE ARM ARMOR STATS
		a = wt = 0
		for i in arm:
			a += i[0]
			wt += i[1]
		try:
			arm_armor += a / len(arm)
			total_weight += wt / len(arm)
		except ZeroDivisionError:
			pass

		# AVERAGE LEG ARMOR STATS
		l = wt = 0
		for i in leg:
			l += i[0]
			wt += i[1]
		try:
			leg_armor += l / len(leg)
			total_weight += wt / len(leg)
		except ZeroDivisionError:
			pass


		print(name)
		print(head)
		print(shld)
		print(body)
		print(arm)
		print(leg)
		print(head_armor, body_armor, arm_armor, leg_armor, total_weight)
		print()

		troop_2D_list.append([name, head_armor, body_armor, arm_armor, leg_armor, total_weight])

	df = pd.DataFrame(troop_2D_list)

	with ExcelWriter('MB2.xlsx', mode='a') as writer:
		df.to_excel(writer, sheet_name='npcAvgArmor', index=False)


if __name__ == '__main__':
	# get_armor('head_ar.xml', 1, 'head_armor', None, None)
	# get_armor('shoulder_ar.xml', 2, 'body_armor', 'arm_armor', None)
	# get_armor('body_ar.xml', 3, 'body_armor', 'arm_armor', 'leg_armor')
	# get_armor('arm_ar.xml', 1, 'arm_armor', None, None)
	# get_armor('leg_ar.xml', 1, 'leg_armor', None, None)
	get_npc('npcs.xml')
	convert_id_to_name('all_equipment.xml')
	get_npc_armor_avg()

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl import Workbook, load_workbook
import pandas as pd
from pandas import ExcelWriter
import csv
import numpy as np


driver = webdriver.Chrome()
driver.implicitly_wait(10)


def troop_skills():
	
def main():
	driver = webdriver.Chrome()
	driver.implicitly_wait(10)

	df = pd.read_excel("MB2.xlsx", sheet_name="npcRevised")
	beta_ver = "1.5.10 beta"

	for index, row in df.iterrows():
		for columnIndex, value in row.items():
			

	https://mountandblade.fandom.com/wiki/Aserai_Recruit?veaction=editsource


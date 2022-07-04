#!/usr/bin/env python
# coding: utf-8

# Import Libraries
from selenium import webdriver
from selenium.webdriver.common.keys import Keys as k
from time import sleep as time
import xlsxwriter as xwr
import os

# Variables
value_dollar = 0
value_euro = 0

# Opened browser
browser = webdriver.Chrome("/home/julio_gabriel/Meus Arquivos/Meus arquivos/Estudo/Cursos/Python/Curso_RPA_Python/Programs Created/chromedriver")

# Openig the browser
browser.get('https://www.google.com.br/')
time(2)

# Camp for find
camp_search = ('q')

# Clicked and search the values
browser.find_element('name',camp_search).send_keys('Whats the value of dolar today in Brazil')
time(1)
browser.find_element('name', camp_search).send_keys(k.RETURN)
time(1)

# looking for the class that contains the dollar value
value_dollar = browser.find_elements(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

# Camp for find
camp_search = ('q')

# Opening new tab
browser.switch_to.new_window('tab')
browser.get('https://www.google.com.br/')

# Clicked and search the values
browser.find_element('name',camp_search).send_keys('Whats the value of euro today in Brazil')
time(1)
browser.find_element('name', camp_search).send_keys(k.RETURN)
time(1)

# looking for the class that contains the euro value
value_euro = browser.find_elements(
    'xpath', '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text

# Creating the worksheet
path_of_file = ('/home/julio_gabriel/Meus Arquivos/Meus arquivos/Estudo/Cursos/Python/Curso_RPA_Python/Programs Created/Programa para abrir o google e perquisar o valor do dolar e do euro/values.xlsx')
file = xwr.Workbook(path_of_file)
sheet1 = file.add_worksheet()

# Passing values
sheet1.write("A1", "Nome da Moeda")
sheet1.write("B1", "Valor da Moeda")
sheet1.write("A2", "Dolar")
sheet1.write("A3", "Euro")
sheet1.write("B2", value_dollar)
sheet1.write("B3", value_euro)
file.close()
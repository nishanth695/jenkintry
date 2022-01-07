from selenium import webdriver;
from selenium.webdriver.common.keys import Keys; 
import time
from selenium.webdriver.common.action_chains import ActionChains
import unittest
import logging
from random import randint

from selenium.common.exceptions import NoSuchElementException
driver = webdriver.Chrome(executable_path='./drivers/chromedriver')

driver.get("http://127.0.0.1:8000/login")
driver.maximize_window()
username = driver.find_element_by_name("emailid")
text = "vinay.gat1@gmail.com"
for character in text:
    username.send_keys(character)
    time.sleep(0.3)

password = driver.find_element_by_name("pwd")
text = "superman"
for character in text:
    password.send_keys(character)
    time.sleep(0.3)



submit = driver.find_element_by_class_name('btn')
submit.click()

alert_obj = driver.switch_to.alert
alert_obj.accept()






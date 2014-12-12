from pyimb import imb
import model
import logging
import os.path
import json
import time

url = 'imb.lohman-solutions.com'
port = 4000
owner_id = 123
owner_name = 'SP'
federation = 'ecodistrict'

try:
    c = imb.Client(url, port, owner_id, owner_name, federation)
    m = model.RenobuildModel()
    m.client = c

    time.sleep(0.5)
    input('Listening for input... Press return to stop module.')

finally:
    c.disconnect()
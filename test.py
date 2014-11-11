from pyimb import imb
import module
import logging
import os.path
owner_id = 123
owner_name = 'rasmus'
federation = 'ecodistrict'

c = imb.Client(imb.TEST_URL, imb.TEST_PORT, owner_id, owner_name, federation)

try:
    m = module.ExcelModule()
    m.client = c

    input()
    
finally:
    c.disconnect()
#print(c.unique_client_id)
# models = c.subscribe('models')
# dashboard = c.publish('dashboard')

# #input()
# #anything.publish()
# #input()

# response = ('{"method": "getModels", "type": "response", "name": "Module name", "id": "moduleId",'
#     '"description": "Description about the module", "kpiList": ["kpi1"]}')

# models.add_handler(imb.ekNormalEvent, 
#     lambda payload: print('NormalEvent', payload))

# input()
# dashboard.signal_string(response)

# input()
# c.disconnect()

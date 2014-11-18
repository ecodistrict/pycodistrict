from pyimb import imb
import model
import logging
import os.path
import json
owner_id = 123
owner_name = 'rasmus'
federation = 'ecodistrict'

#c = imb.Client(imb.TEST_URL, imb.TEST_PORT, owner_id, owner_name, federation)

try:
    m = model.RenobuildModel()
    #m.client = c

    request = {
    "method": "getModels",
    "type": "request",
    "parameters": {
        "kpiList": "kpi1"
    }
    }

    #m._handle_request(imb.encode_string(json.dumps(request)))

    inputs = [
        {
            "id": "time-frame",
            "type": "number",
            "value": 50
        },
        {
            "id": "buildings",
            "type": "list",
            "inputs": [
                [
                    {
                        "id": "building-name",
                        "type": "text",
                        "value": "building A"
                    },
                    {
                        "id": "heating-system",
                        "type": "select",
                        "value": 1
                    },
                    {
                        "id": "energy-use",
                        "type": "number",
                        "value": 5000
                    }
                ],
                [
                    {
                        "id": "building-name",
                        "type": "text",
                        "value": "building A"
                    },
                    {
                        "id": "heating-system",
                        "type": "select",
                        "value": 2
                    },
                    {
                        "id": "energy-use",
                        "type": "number",
                        "value": 10000
                    }
                ]
            ]
        }
    ]

    request = {
        "method": "startModel",
        "type": "request",
        "moduleId": "sp-renobuild-excel-model",
        "variantId": "variantId",
        "kpiAlias": "energy-kpi",
        "inputs": inputs
    }

    #m._handle_request(imb.encode_string(json.dumps(request)))
    print(m.run_model(inputs, 'energy-kpi'))

    #input()

finally:
    pass#c.disconnect()
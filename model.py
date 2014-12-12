import uuid
import json
from pyimb import imb
import logging
from enum import Enum
import threading
from functools import partial
import win32com.client
from pywintypes import com_error
import os.path
import statistics

MODELS_EVENT = 'models'
DASHBOARD_EVENT = 'dashboard'

class ModelStatus(Enum):
    """Statuses of a Model"""
    PROCESSING = 'processing'
    SUCCESS = 'success'

class Model(object):
    """ECODISTR-ICT model"""
    def __init__(self,):
        super().__init__()

    @property
    def name(self):
        return self._name
    
    @property
    def kpi_list(self):
        return self._kpi_list

    @property
    def id(self):
        return self._id
    
    @property
    def description(self):
        return self._description

    @property
    def input_specification(self):
        return self._input_specification

    @property
    def client(self):
        return self._client
    @client.setter
    def client(self, value):
        self._client = value
        logging.debug('Registering client...')
        models_event = self._client.subscribe('models')
        logging.debug('Subscribed to models')
        models_event.add_handler(imb.ekNormalEvent, self._handle_request)
        self._dashboard_event = self._client.publish('dashboard')
        logging.debug('Published to dashboard')

    def _send_status(self, status, request):
        response = {'method': request['method'], 'type': 'response'}
        response['moduleId'] = self.id
        response['variantId'] = request['variantId']
        response['kpiAlias'] = request['kpiAlias']
        response['status'] = status.value
        self._send_message(response)

    def _handle_request(self, payload):
        request = json.loads(imb.decode_string(payload))

        assert(request['type'] == 'request')

        logging.debug('Got request: {0}'.format(request['method']))

        response = {'method': request['method'], 'type': 'response'}

        if request['method'] == 'getModels':
            response['name'] = self.name
            response['id'] = self.id
            response['description'] = self.description
            response['kpiList'] = self.kpi_list
            self._send_message(response)

        elif request['method'] == 'selectModel':
            if request['moduleId'] != self.id:
                return
            response['moduleId'] = self.id
            response['variantId'] = request['variantId']
            response['kpiAlias'] = request['kpiAlias']
            response['inputs'] = self._input_specification
            self._send_message(response)

        elif request['method'] == 'startModel':
            if request['moduleId'] != self.id:
                return            

            self._send_status(ModelStatus.PROCESSING, request)
            t = threading.Thread(target=partial(self._run_and_respond, request))
            t.start()


        else:
            raise NotImplementedError('Method {0} not implemented.'.format(request[method]))

    def _send_message(self, message):
        logging.debug('Sending message: {0}'.format(message))

        self._dashboard_event.signal_string(json.dumps(message))

    def _run_and_respond(self, request):
        model_inputs = request['inputs']
        kpi_alias = request['kpiAlias']
        model_outputs = self.run_model(model_inputs, kpi_alias)

        message = {
            'method': 'modelResult',
            'type': 'result',
            'moduleId': self.id,
            'variantId': request['variantId'],
            'kpiAlias': kpi_alias,
            'outputs': model_outputs
        }

        self._send_message(message)

        self._send_status(ModelStatus.SUCCESS, request)


class RenobuildModel(Model):
    """docstring for RenobuildModel"""
    def __init__(self):
        super().__init__()
        self._name = "SP Renobuild LCA tool"
        self._id = "sp-renobuild-excel-model"
        self._kpi_list = ['energy-kpi', 'ghg-kpi']
        self._description = "Renobuild is SP's LCA-based tool for assessing building renovations."
        self._cell_addresses = {
            'time-frame': ('renobuild-test.xlsx', 'Calc', (15, 3)),
            'heating-system': ('renobuild-test.xlsx', 'Calc', (92, 3)),
            'energy-use': ('renobuild-test.xlsx', 'Calc', (93, 3)),
            'energy-kpi': ('renobuild-test.xlsx', 'ECODISTRICT', (1, 2)),
            'ghg-kpi': ('renobuild-test.xlsx', 'ECODISTRICT', (2, 2))
        }

        self._input_specification = json.loads("""
            [
                {
                    "id": "time-frame",
                    "type": "number",
                    "label": "Time period for LCA calculation",
                    "unit": "years",
                    "min": 1,
                    "value": 50
                },
                {
                    "type": "list",
                    "label": "Define buildings and their heating systems",
                    "id": "buildings",
                    "inputs": [
                        {
                            "id": "name",
                            "type": "text",
                            "label": "Building name"
                        },
                        {
                            "id": "heating-system",
                            "label": "Heating system",
                            "type": "select",
                            "options": [{
                              "id": "4",
                              "label": "Individual gas boilers"
                            }, {
                              "id": "2",
                              "label": "District heating"
                            }, {
                              "id": "1",
                              "label": "Individual heat pums"
                            }],
                            "value": "individual-gas-boilers"
                        },
                        {
                            "id": "energy-use",
                            "type": "number",
                            "label": "Annual heat energy use",
                            "unit": "kWh/year",
                            "min": 0,
                            "value": 5000
                        }
                    ]
                }
            ]
            """)

    def make_input_data_dict(self, inputs):
        return {inp['id']: inp['value'] for inp in inputs}

    def run_model(self, inputs, kpi_alias):
        inputs = self.make_input_data_dict(inputs)

        # Initialize COM in this thread
        win32com.client.pythoncom.CoInitialize()
        excel = win32com.client.Dispatch('Excel.Application')
        workbooks = {}
        try:

            file_paths = set((address[0] for address in self._cell_addresses.values()))
            sheets = {}
            cells = {}
            for key, addr in self._cell_addresses.items():
                if not addr[0] in workbooks:
                    workbooks[addr[0]] = excel.Workbooks.Open(os.path.abspath(addr[0]))
                wb = workbooks[addr[0]]
                if not (addr[0], addr[1]) in sheets:
                    sheets[(addr[0], addr[1])] = wb.Worksheets(addr[1])
                sheet = sheets[(addr[0], addr[1])]
                cells[key] = sheet.Cells(*addr[2])
            
            cells['time-frame'].Value = inputs['time-frame']

            def compute_building_kpi(building):
                cells['heating-system'].Value = building['heating-system']
                cells['energy-use'].Value = building['energy-use']
                return cells[kpi_alias].Value

            outputs = []

            kpi_list = [
                {'name': b['name'], 'kpiValue': compute_building_kpi(b)}
                for b in inputs['buildings']]

            outputs.append({
                'type': 'kpi',
                'info': 'Average value of KPI {0}'.format(kpi_alias),
                'value': int(statistics.mean((b['kpiValue'] for b in kpi_list))),
                'unit': 'kWh/year'
                })

            outputs.append({
                'type': 'kpi-list',
                'label': 'KPI {0} for each building'.format(kpi_alias),
                'value': kpi_list,
                'unit': 'kWh/year'
                })

        finally:
            for wb in workbooks.values():
                wb.Close(False)
            excel.Quit()
            win32com.client.pythoncom.CoUninitialize()

        return outputs

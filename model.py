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

MODELS_EVENT = 'models'
DASHBOARD_EVENT = 'dashboard'

class ModelStatus(Enum):
    """Statuses of a Model"""
    STARTING = 'starting'

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
    def inputs(self):
        return self._inputs    

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

    def _handle_request(self, payload):
        request = json.loads(imb.decode_string(payload))

        assert(request['type'] == 'request')

        logging.debug('Got request: {0}'.format(request))

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
            response['inputs'] = self._inputs
            self._send_message(response)

        elif request['method'] == 'startModel':
            if request['moduleId'] != self.id:
                return
            response['moduleId'] = self.id
            response['variantId'] = request['variantId']
            response['kpiAlias'] = request['kpiAlias']
            response['status'] = ModelStatus.STARTING.value
            self._send_message(response)

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

class RenobuildModel(Model):
    """docstring for RenobuildModel"""
    def __init__(self):
        super().__init__()
        self._name = "Renobuild"
        self._id = "sp-renobuild-excel-model"
        self._kpi_list = ['kpi1', 'kpi2']
        self._description = "Interface to SP's LCA tool Renobuild."
        self._input_cells = {
            'num1': ('test.xlsx', 'Blad1', (3, 3)),
            'num2': ('test.xlsx', 'Blad1', (2, 3))
        }
        self._kpi_cells = {
            'kpi1': ('test.xlsx', 'Blad2', (1, 1)),
            'kpi2': ('test.xlsx', 'Blad2', (1, 2))
        }

        self._inputs = [
            {
                "type": "input-group",
                "label": "Some inputs",
                "inputs": [
                    {
                        "label": "A number (cell C3)",
                        "type": "number",
                        "id": "num1",
                        "unit": "m",
                        "min": 50,
                        "max": 100
                    },
                    {
                        "label": "Another number (cell C2)",
                        "type": "number",
                        "min": 0,
                        "unit": "m",
                        "digits": 2,
                        "id": "num2"
                    }
                ]
            }
        ]

    def run_model(self, inputs, kpi_alias):
        # Initialize COM in this thread
        win32com.client.pythoncom.CoInitialize()
        excel = win32com.client.Dispatch('Excel.Application')
        all_cells = tuple(self._input_cells.values()) + tuple(self._kpi_cells.values())
        all_paths = set([c[0] for c in all_cells])
        all_paths = {key: os.path.abspath(key) for key in all_paths}

        workbooks = {key: excel.Workbooks.Open(path) for (key, path) in all_paths.items()}
        for input_id, value in inputs.items():
            filename, sheet, coords = self._input_cells[input_id]
            cell = workbooks[filename].Worksheets(sheet).Cells(*coords)
            cell.Value = value

        filename, sheet, coords = self._kpi_cells[kpi_alias]
        kpi_cell = workbooks[filename].Worksheets(sheet).Cells(*coords)
        kpi_value = kpi_cell.Value

        for wb in workbooks.values():
            wb.Close(False)

        return kpi_value
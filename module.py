import uuid
import json
from pyimb import imb
import logging
from enum import Enum
import threading
from functools import partial

MODELS_EVENT = 'models'
DASHBOARD_EVENT = 'dashboard'

class ModelStatus(Enum):
    """Statuses of a Model"""
    STARTING = 'starting'

class Module(object):
    """ECODISTR-ICT module"""
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

    def _run_and_respond(request):
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

class ExcelModule(Module):
    """docstring for ExcelModule"""
    def __init__(self):
        super().__init__()
        self._name = "Excel module"
        self._id = "sp-excel-module"
        self._description = "Interface to an Excel file."
        self._kpi_list = ['kpi1', 'kpi2']

    def run_model(self, inputs, kpi_alias):
        return 'ExcelModule calculated KPI {0}!'.format(kpi_alias)

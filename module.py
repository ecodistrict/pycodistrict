import uuid
import json
from pyimb import imb
import logging

MODELS_EVENT = 'models'
DASHBOARD_EVENT = 'dashboard'

class Module(object):
    """ECODISTR-ICT module"""
    def __init__(self,):
        super().__init__()

    @property
    def name(self):
        return self._name
    
    @property
    def kpilist(self):
        return self._kpilist

    @property
    def id(self):
        return self._id
    
    @property
    def description(self):
        return self._description

    @property
    def client(self):
        return self._client
    @client.setter
    def client(self, value):
        self._client = value
        models_event = self._client.subscribe('models')
        logging.debug('Subscribed to models')
        models_event.add_handler(imb.ekNormalEvent, self._handle_request)
        self._dashboard_event = self._client.publish('dashboard')
        logging.debug('Published to dashboard')

    def _handle_request(self, payload):
        request = json.loads(imb.decode_string(payload))

        logging.debug('Got request: {0}'.format(request))

        response = {'method': request['method'], 'type': 'response'}

        if request['method'] == 'getModels':
            response['name'] = self.name
            response['id'] = self.id
            response['description'] = self.description
            response['kpilist'] = self.kpilist
            self._send_message(response)
        else:
            raise NotImplementedError()

    def _send_message(self, message):
        logging.debug('Sending message: {0}'.format(message))
        self._dashboard_event.signal_string(json.dumps(message))
    

class ExcelModule(Module):
    """docstring for ExcelModule"""
    def __init__(self):
        super().__init__()
        self._name = "Excel module"
        self._id = "sp-excel-module"
        self._description = "Interface to an Excel file."
        self._kpilist = ['kpi1', 'kpi2']


    
    
    
    
    
    
    
    
    
    
        
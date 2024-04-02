import os
import time
import subprocess
import win32com.client

class client:
    '''
    SAP client interface for python
    '''
    def __init__(self, login:str=None) -> None:
        self.session = None
        self.connection = None
        self.application = None
        self.SapGuiAuto = None
        if login is not None: self.login(login)
    
    def _field_selector(self, query):
        s = ''
        if 'wnd' not in query: s += 'wnd[0]/'
        if 'usr' not in query: s += 'usr/'
        s += query
        return s
    
    def login(self,
              sap_server:str,
              sap_path:str="C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe") -> None:
        '''
        Login to a specified SAP server.

        Parameters:
            sap_server:         SAP server string (as listed in the window)
            sap_path:           Path to SAP gui program file
        '''
        sap = subprocess.Popen(sap_path)
        time.sleep(1)

        self.SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(self.SapGuiAuto) == win32com.client.CDispatch: raise ValueError(f"SAP GUI was not found!")

        self.application = self.SapGuiAuto.GetScriptingEngine
        if not type(self.application) == win32com.client.CDispatch: raise ValueError(f"SAP Scripting Engine was not found!")

        if self.application.Connections.Count >= 2: raise ValueError("Too many pre-existing SAP connections, please close all but 1 SAP connections!")
        elif 2 > self.application.Connections.Count > 0: self.connection = self.application.Children(0)
        else: self.connection = self.application.OpenConnection(sap_server, True)
        
        if not type(self.connection) == win32com.client.CDispatch: raise ValueError("Could not establish connection to SAP server!")

        self.session = self.connection.Children(0)
        if not type(self.session) == win32com.client.CDispatch: raise ValueError("Could not attach to open connection to SAP server!")
   
    def logout(self) -> None:
        '''
        De-attach python from the SAP client and try to close the conection
        '''
        if self.session is not None: self.open_transaction("/nex"); self.session = None
        if self.connection is not None: self.connection = None
        if self.application is not None: self.application = None
        if self.SapGuiAuto is not None: self.SapGuiAuto = None

    def open_transaction(self, transaction):
        '''
        Open the specified transaction window
        '''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = f"{transaction}"
            self.session.findById("wnd[0]").sendVKey(0)
        except Exception as e:
            raise ValueError(f"{transaction} could either not be found or there is a problem with the SAP connection, more details:\n{e}")

    def find_elements(self, path):
        '''
        Recursively find all elements that contain the specified text/path.
        Path of the elements can be accessed with element.Id and the text inside of the element with element.text

        Parameters:
            element:    The current element to search
            path:       The path to search for
        '''
        def search_elements(element, path):
            result = []
            if hasattr(element, 'Children'):
                try:
                    for i in range(element.Children.Count):
                        child = element.Children.ElementAt(i)
                        if path in child.Id:
                            result.append(child)
                        result.extend(search_elements(child, path))
                except: pass
            return result
        return search_elements(self.session, path)
    
    def update_field(self, field, value):
        '''
        Set field(s) to specified values.

        Parameters:
            field:              Id of the input field(s)
            value:              Text to set on the input field(s)
        '''
        if isinstance(field, str): field = field.split()
        if isinstance(value, str): value = value.split()
        if len(field) != len(value): raise ValueError("Provide the same number of fields and values!")

        for f, v in zip(field,value):
            self.session.findById(self._field_selector(f)).text = v

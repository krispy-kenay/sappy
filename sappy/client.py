import os
import json
import time
import subprocess
import win32com.client

class Client:
    '''
    SAP client interface for python
    '''
    def __init__(self, login:str=None) -> None:
        self.session = None
        self.connection = None
        self.application = None
        self.sap_gui_auto = None
        if login is not None:
            self.login(login)

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

        self.sap_gui_auto = win32com.client.GetObject('SAPGUI')
        if not isinstance(self.sap_gui_auto, win32com.client.CDispatch):
            raise ValueError("SAP GUI was not found!")

        self.application = self.sap_gui_auto.GetScriptingEngine
        if not isinstance(self.application, win32com.client.CDispatch):
            raise ValueError("SAP Scripting Engine was not found!")

        if self.application.Connections.Count >= 2:
            raise ValueError("Too many pre-existing SAP connections, please close all but 1 SAP connections!")
        if 2 > self.application.Connections.Count > 0:
            self.connection = self.application.Children(0)
        else: self.connection = self.application.OpenConnection(sap_server, True)

        if not isinstance(self.connection, win32com.client.CDispatch):
            raise ValueError("Could not establish connection to SAP server!")

        self.session = self.connection.Children(0)
        if not isinstance(self.session, win32com.client.CDispatch):
            raise ValueError("Could not attach to open connection to SAP server!")
        self.open_transaction('/n')

    def logout(self) -> None:
        '''
        De-attach python from the SAP client and try to close the conection
        '''
        if self.session is not None:
            self.open_transaction("/nex")
            self.session = None
        if self.connection is not None:
            self.connection = None
        if self.application is not None:
            self.application = None
        if self.sap_gui_auto is not None:
            self.sap_gui_auto = None

    def open_transaction(self, transaction:str) -> None:
        '''
        Open the specified transaction window
        '''
        try:
            self.session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            self.send_key(0)
            self.session.findById("wnd[0]/tbar[0]/okcd").text = f"{transaction}"
            self.send_key(0)
        except Exception as e:
            raise ValueError(f"{transaction} could either not be found or there is a problem with the SAP connection, more details:\n{e}") from e

    def send_key(self, key:int|list|tuple|set, window:int=0)-> None:
        '''
        Send specified key(s) to the specified window
        '''
        if isinstance(key, int):
            key = [key]

        for k in key:
            self.session.findById(f"wnd[{window}]").sendVKey(k)

    def find_elements(self, idn:str) -> list[str]:
        '''
        Find all elements that contain the specified text/Id using the GetObjectTree method.

        Parameters:
            Id:                 The Id to search for
        Returns:
            results:            List with paths to all matching elements
        '''
        object_tree = json.loads(self.session.GetObjectTree('wnd[0]/usr/'))
        def search_tree(tree):
            result = []
            if idn in tree.get('properties',{}).get('Id',''):
                result.append(tree['properties']['Id'])
            for child in tree.get('children', []):
                result.extend(search_tree(child))
            return result
        return search_tree(object_tree)

    def find_element(self, idn:str) -> object:
        '''
        Wrapper for find_elements where the search Id is expected to be unique.
        This should make accessing elements by Id easier, since the entire path is no longer needed.

        Parameters:
            Id:                 The Id to search for
        '''
        elements = self.find_elements(idn)
        if not elements:
            raise ValueError(f"No element found with Id: {idn} in its path")
        if len(elements) > 1:
            raise ValueError(f"More than one element found with Id: {idn} in its path")
        return self.session.findById(elements[0])

    def update_field(self, idn:str|list|tuple|set, value:str|list|tuple|set) -> None:
        '''
        Set field(s) to specified values.

        Parameters:
            idn:                Id of the input field(s)
            value:              Text to set on the input field(s)
        '''
        if isinstance(idn, str):
            idn = idn.split()
        if isinstance(value, str):
            value = value.split()
        if len(idn) != len(value):
            raise ValueError("Provide the same number of fields and values!")

        for f, v in zip(idn,value):
            self.find_element(f).text = v

    def get_table(self, idn:str) -> list:
        '''
        Return table in SAP as a list (only for GuiTableControl Types at the moment)

        Parameters:
            idn:                Id of the table
        Returns:
            lis:                List content in SAP
        '''
        if not isinstance(idn, str):
            raise ValueError("Provide Id to table as a string!")
        table = self.find_element(idn)
        if not table.Type == 'GuiTableControl':
            raise TypeError("The element needs to be a 'GuiTableControl' type object! Check what type it is with element.Type")

        li = []
        # Get number of columns
        columns = 0
        while True:
            try:
                table.GetCell(0,columns)
                columns += 1
            except Exception:
                break

        # Get table data
        for i in range(table.VisibleRowCount):
            subl = []
            for j in range(columns):
                try:
                    subl.append(table.GetCell(i,j).text)
                except Exception:
                    pass
            li.append(subl)
        return li

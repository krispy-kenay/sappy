import json
import time
import subprocess
import win32com.client

class Client:
    '''
    SAP Client interface in Python
    '''
    def __init__(self, sap_path:str="C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe") -> None:
        self.process = None
        self.connection = None
        self.application = None
        self.sap_gui_auto = None
        self._attach_to_engine(sap_path)
    
    def _attach_to_engine(self, sap_path) -> None:
        '''
        Attach to SAP scripting engine

        Parameters:
            sap_path:           Path to SAP executable
        '''
        self.process = subprocess.Popen(sap_path)

        self.sap_gui_auto = win32com.client.GetObject('SAPGUI')
        if not isinstance(self.sap_gui_auto, win32com.client.CDispatch):
            raise ValueError("SAP GUI was not found!")

        self.application = self.sap_gui_auto.GetScriptingEngine
        if not isinstance(self.application, win32com.client.CDispatch):
            raise ValueError("SAP Scripting Engine was not found!")

    def _open_connection(self, sap_server:str) -> bool:
        '''
        Open a new connection in SAP or attach to an existing one.
        Returns a boolean to indicate whether a pre-existing valid connection was found.

        Parameters:
            sap_server:         SAP server id
        '''
        count = self.application.Connections.Count
        if count == 0:
            self.connection = self.application.OpenConnection(sap_server, True)
        else:
            for i in range(count):
                if self.application.Children(i).Description == sap_server:
                    self.connection = self.application.Children(i)
                    print("Attached to existing connection")
                    if not isinstance(self.connection, win32com.client.CDispatch):
                        raise ValueError(f"Connection to server: {sap_server} could not be established!")
                    return True
            print("Opening new connection")
            self.connection = self.application.OpenConnection(sap_server, True)

        if not isinstance(self.connection, win32com.client.CDispatch):
            raise ValueError(f"Connection to server: {sap_server} could not be established!")

        return False

    def new_session(self, sap_server:str) -> object:
        '''
        Return a new SAP session, intended to be used with a context manager (like "with")

        Parameters:
            sap_server:         SAP server id
        '''
        if not self.connection or self.connection.Description != sap_server:
            was_open = self._open_connection(sap_server)
            master_session = self.connection.Children(0)
            if not was_open:
                print("Connection not found")
                return Client.Session(master_session)    
        else:         
            master_session = self.connection.Children(0)
        
        before = set([child.Id for child in self.connection.Children])
        master_session.createSession()
        
        after = set([child.Id for child in self.connection.Children])
        start = time.time()
        current = time.time()

        while not(after - before) and (current - start) < 20:
            current = time.time()
            after = set([child.Id for child in self.connection.Children])

        selected_child = list(after - before)[0]
        session = self.connection.findById(selected_child)

        return Client.Session(session)

    class Session:
        '''
        SAP GUI instance/window handler
        '''
        def __init__(self, ses:win32com.client.CDispatch) -> None:
            self.ses = ses
        
        def __enter__(self) -> object:
            return self

        def __exit__(self, exc_type, exc_val, exc_tb) -> None:
            self.close()
        
        def close(self) -> None:
            '''
            Close the session. After this it can no longer be used to perform actions!
            '''
            self.ses.findById("wnd[0]").close()
        
        def open_transaction(self, transaction:str) -> None:
            '''
            Open the specified transaction id

            Parameters:
                transaction:        Transaction id to open in current session
            '''
            try:
                self.close_transaction()
                self.ses.findById("wnd[0]/tbar[0]/okcd").text = f"{transaction}"
                self.send_key(0)
            except Exception as e:
                raise ValueError(f"{transaction} could either not be found or there is a problem with the SAP connection, more details:\n{e}") from e
        
        def close_transaction(self) -> None:
            '''
            Close the currently open transaction
            '''
            try:
                self.ses.findById("wnd[0]/tbar[0]/okcd").text = "/n"
                self.send_key(0)
            except Exception as e:
                raise ValueError(f"Unable to close connection due to:\n{e}") from e
        
        def send_key(self, key:int|list|tuple|set, window:int=0)-> None:
            '''
            Send a key input to the session/window

            Parameters:
                key:                Key id to send (see here: https://experience.sap.com/files/guidelines/References/nv_fkeys_ref2_e.htm)
                window:             Id of the window, defaults to 0 (main window)
            '''
            if isinstance(key, int):
                key = [key]

            for k in key:
                self.ses.findById(f"wnd[{window}]").sendVKey(k)

        def find_elements(self, idn:str) -> list[str]:
            '''
            Find all elements that contain the specified text/Id using the GetObjectTree method.

            Parameters:
                idn:                Id of the elements to search for
            Returns:
                results:            List with paths to all matching elements
            '''
            object_tree = json.loads(self.ses.GetObjectTree(''))
            def search_tree(tree):
                result = []
                if idn in tree.get('properties',{}).get('Id',''):
                    result.append(tree['properties']['Id'])
                for child in tree.get('children', []):
                    result.extend(search_tree(child))
                return result
            return search_tree(object_tree)

        def find_element(self, idn:str, first_element:bool=False) -> object:
            '''
            Wrapper for find_elements where the search Id is expected to be unique.
            This should make accessing elements by Id easier, since the entire path is no longer needed.

            Parameters:
                idn:                Id of the element to search for
                first_element:      Override if first item should be taken always (even if there are multiple)
            Returns:
                element:            Found element as a reference to the object
            '''
            elements = self.find_elements(idn)
            if not elements:
                raise ValueError(f"No element found with Id: {idn} in its path")
            if len(elements) > 1 and not first_element:
                raise ValueError(f"More than one element found with Id: {idn} in its path")
            return self.ses.findById(elements[0])
        
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
            Return table in SAP as a list (for GuiTableControl and GridViewCtrl for now)

            Parameters:
                idn:                Id of the table element
            Returns:
                output:             Content of the Table from SAP
            '''
            if not isinstance(idn, str):
                raise ValueError("Provide Id to table as a string!")
            table = self.find_element(idn, first_element=True)

            def GuiTableControl(element):
                output = []
                for row in range(element.RowCount):
                    content, column = [], 0
                    while True:
                        try:
                            content.append(element.GetCell(row, column).text)
                        except: break
                        column += 1
                    output.append(content)
                return output

            def GridViewCtrl(element):
                output = []
                for row in range(element.RowCount):
                    if row % 3 == 0: element.firstVisibleRow = str(row)
                    content = []
                    for j,column in enumerate(element.ColumnOrder):
                        try:
                            if j % 3 == 0: element.firstVisibleColumn = column
                            content.append(element.getcellvalue(row,column))
                        except: pass
                    output.append(content)
                return output

            match table.Type:
                case 'GuiTableControl':
                    output = GuiTableControl(table)
                case 'GuiShell':
                    if 'GridViewCtrl' in table.Text:
                        output = GridViewCtrl(table)
                    else:
                        raise TypeError(f"{table.Text} type object is not supported!")
                case _:
                    raise TypeError(f"{table.Type} type object is not supported!")

            output = [row for row in output if row]
            return output
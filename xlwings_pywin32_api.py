"""
/// @file xlwings_pywin32_api.py
/// @author Austin Vandegriffe
/// @date 2020-05-20
/// @brief A WORK IN PROGRESS: Excel handle for Python.
/// ## Goal: Better VBA communication.
/// @pre N/A
/// @style K&R, and "one true brace style" (OTBS), and '_' variable naming
/////////////////////////////////////////////////////////////////////
/// @references
/// ## [1] https://github.com/mhammond/pywin32
/// ## [2] https://docs.xlwings.org/en/stable/
/// ## [3] https://stackoverflow.com/a/28173146
"""

import xlwings as xw

class XLWings_PyWin32_Handle(object):
    __XLWings_PyWin32_Handle_zero_indexed = True
    def __init__(self, app : xw.main.App):
        self.__app : xw.App = app
        self.__app.Visible = True 
        self.__app.DisplayAlerts = False
        self.__wb = None
        self.worksheets = {}
        self.vba_modules = {}
        self.active_sheet = None
        self.cell = None
        self.range = None
        
    def get_api(self):
        """
            The api is private and cannot be accessed.
        """
        raise NotImplementedError('The api is private and cannot be accessed.')
    def get_workbook(self):
        raise NotImplementedError('The workbook is private to prevent users from ' +
                                    'improperly reasigning the workbook pointer.')
    
#     def sync_active_xlwings(self):
#         xw.apps[self.api.pid]
        
    def create_workbook(self):
        """
            Create a new workbook with 1 default sheet ('Sheet1') and a VBA module.
        """
        if self.__wb == None:
            self.__wb = self.__app.api.Workbooks.Add()
            self.worksheets = {sh.Name:sh for sh in self.__wb.Sheets}
            self.vba_modules = {m.Name:m for m in self.__wb.VBProject.VBComponents}
            # self.__wb.FullName = ""
        else:
            raise Exception('XLWings_PyWin32_Handle only handles ONE notebook at a time. ' +
                                'Create a new instance for second notebook.')
    
    def load_workbook(self, fullname):
        """
            Load a workbook and all its worksheets and VBA modules.
        """
        if self.__wb == None:
            self.__wb = self.__app.api.Workbooks.Open(f"{fullname}")
            self.worksheets = {sh.Name:sh for sh in self.__wb.Sheets}
            self.vba_modules = {m.Name:m for m in self.__wb.VBProject.VBComponents}
        else:
            raise Exception('XLWings_PyWin32_Handle only handles ONE notebook at a time. ' +
                                'Create a new instance for second notebook.')
    
    def create_worksheet(self, name=None):
        """
            Create a new, blank worksheet. See [3].
        """
        t_ws = self.__wb.Worksheets.Add()
        if name is None:
            name = t_ws.Name
        t_ws.Name = name
        self.worksheets[name] = t_ws
        t_ws = None
        
    def activate_worksheet(self, name):
        """
            Activate a specific worksheet, IT MUST ALREADY EXISTS! 
            If it doesn't, create it with "create_worksheet".
        """
        if name in self.worksheets:
            self.activate_worksheet = self.worksheets[name]
            self.cell = self.activate_worksheet.Cells
            self.range = self.activate_worksheet.Range
        else:
            raise KeyError(f'Worksheet "{name}" does not exists. Either ' +
                                f'create it or choose from {self.worksheets.keys}')
    
    def add_vba_module(self, module_name = None):
        """
            Add a code module to active workbook. See [3].
        """
        if module_name in self.vba_modules:
            raise Exception(f'Module "{module_name}" already exists. ' +
                                'Choose a different name.')
        t_module = self.__wb.VBProject.VBComponents.Add(1)
        if module_name:
            t_module.Name = module_name
        self.vba_modules[t_module.Name] = t_module
    
    def activate_vba_module(self, module_name):
        """
            Choose a VBA module for writing VBA to in "add_vba".
        """
        if module_name in self.vba_modules:
            self.active_module = self.vba_modules[module_name]
        else: 
            KeyError(f'Modules "{module_name}" does not exists. ' +
                        f'Either create it or choose from {self.vba_modules}')
            
    def add_vba(self, vba_code, module_name=None):
        """
            Add VBA code to active module or specify the module to add to. See [3].
        """
        if module_name:
            self.vba_modules[module_name].CodeModule.AddFromString(vba_code)
        else:
            self.active_module.CodeModule.AddFromString(vba_code)

    def run_vba(self, routine_name, *args):
        """
            Run a macro or function with arguments in *args. See [3].
        """
        return self.__wb.Application.Run(routine_name, *args)
        
    if __XLWings_PyWin32_Handle_zero_indexed:
        def __getitem__(self, idx):
            if isinstance(idx, tuple):
                if isinstance(idx[0], tuple):
                    return self.range(
                            self.cell(idx[0][0]+1,idx[0][1]+1),   self.cell(idx[1][0]+1,idx[1][1]+1)
                        )
                else:
                    return self.cell(idx[0]+1,idx[0]+1)
            elif isinstance(idx, str):
                return self.range(idx)
            else:
                raise NotImplementedError('XLWings_PyWin32_Handle does not know ' +
                                            'how to process "{idx}", please implement.')
    else:
        def __getitem__(self, idx):
            if isinstance(idx, tuple):
                if isinstance(idx[0], tuple):
                    return self.range(
                            self.cell(idx[0][0],idx[0][1]),   self.cell(idx[1][0],idx[1][1])
                        )
                else:
                    return self.cell(idx[0],idx[0])
            elif isinstance(idx, str):
                return self.range(idx)
            else:
                raise NotImplementedError('XLWings_PyWin32_Handle does not know how to ' +
                                            f'process "{idx}", please implement.')
    ##############################################################
    # Define class handle for use of "with XLWings_PyWin32_Handle as app: ..."
    def __enter__(self):
        """
            Define initiation of class handle.
        """
        return self
    def __exit__(self, *args) -> None:
        """
            Define exit of class handle.
        """
        self.__del__()
        return
    ##############################################################
    def __del__(self):
        self.__wb.Close(False) # Close without save
        self.__app.quit()


if __name__ == "__main__":

    with XLWings_PyWin32_Handle(xw.App(visible = True, add_book=False)) as wb:
        # THE WORKBOOK AND WORKSHEET
        # Create the workbook
        wb.create_workbook()
        # Create a worksheet
        wb.create_worksheet("Test")
        # Activate the worksheet, i.e. mark it as the one
        ## being reading from and written to.
        wb.activate_worksheet("Test")
        # Write to cell "A1", notice the zero index.
        wb[(0,0)].Value = "Hello World!"

        # Create a VBA module
        wb.add_vba_module("Module1")
        # Activate the module, i.e. mark it as the one
        ## being written to.
        wb.activate_vba_module("Module1")

        macros = {
            "HelloWorld":"""
                            Function HelloWorld() As String
                                MsgBox "Hello from your UDF!", vbOKOnly, "Test"
                                HelloWorld = "UDF execution successful."
                            End Function
                        """
        }

        # Add a user define function (UDF) to the workbook
        wb.add_vba(macros["HelloWorld"])
        # Try out the UDF
        print(wb.run_vba("HelloWorld"))
        
        # input()
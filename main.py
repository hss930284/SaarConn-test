#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### ++++++++++++ ---------- __________ SECTION :  import required Initial Modules and definitions __________ ----------  ++++++++++++  ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

import importlib # Import the importlib module for dynamic module loading
import warnings # Import the warnings module for managing warning messages
import pandas as pd # type: ignore # Import the pandas library for data manipulation and analysis
import openpyxl # type: ignore
import Pkg_struct # Import the Pkg_struct module for package structure definitions
import arelements_def as arelements_def # Import the arelements_def module for AUTOSAR element definitions
import config

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl") # Suppress specific warnings from the openpyxl module

# Reload the arelements_def and Pkg_struct module for dynamic updates
importlib.reload(arelements_def)
importlib.reload(Pkg_struct)

from arelements_def import root # Import the root variable from the arelements_def module

# here, root is dynamic as per the arelements_def module version and it will get selected by user as per AUTOSAR schema version
# for example root for AUTOSAR 4_0_2 schema version, 
# root = ET.Element("AUTOSAR", 
#                      attrib={
#                          "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
#                          "xmlns": "http://autosar.org/schema/r4.0",
#                          "xsi:schemaLocation": "http://autosar.org/schema/r4.0 AUTOSAR_4-0-2.xsd"
#                      })

from collections import defaultdict

from Pkg_struct import ARXMLStructure # Import the ARXMLStructure class from the Pkg_struct module

import rng # Import the rng module for random number generation

import xml.etree.ElementTree as ET # Import the ElementTree class from the xml.etree module

arxml_structure = ARXMLStructure() # Create an instance of the ARXMLStructure class

arxml_structure.create_default_pkg_struct(root) # Create the default package structure using the root element

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### ++++++++++++ ---------- __________ SECTION :  Excel Related Functions __________ ----------  ++++++++++++ ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

from excel_utils import ExcelReader
from data_type_utils import DataProcessor  # Import the DataProcessor class

from itertools import groupby

# Initialize the ExcelReader
excel_reader = ExcelReader()

# Get the file path from the user
excel_reader.get_file_path_from_user()

# Read the Excel file
workbook, excel_file = excel_reader.read_user_defined_excel()

worksheets = {sheet.title: sheet for sheet in workbook.worksheets}
project_info = worksheets['project_info']
swc_info = worksheets['swc_info']
ib_data = worksheets['ib_data']
ports = worksheets['ports']
adt_primitive = worksheets['adt_primitive']
adt_composite = worksheets['adt_composite']
idt = worksheets['idt']

# Process the data
data = excel_reader.read_columns(workbook.active, 'A', 'B')  # Use the active sheet or specify the sheet

# Create an instance of the DataProcessor class
processor = DataProcessor()

# Convert the data types
converted_data = [processor.value_to_str(value) for value in data if isinstance(value, (int, float, bool))]


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  beautify xml __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
def remove_namespaces(element):
    """
    Removes namespace prefixes from element tags.

    Args:
        element: The root element of the XML tree.

    Returns:
        The modified element with namespaces removed.
    """
    for elem in element.iter():
        if '}' in elem.tag:
            elem.tag = elem.tag.split('}', 1)[1]  # Remove namespace prefix
    return element

def indent(elem, level=0):
    """
    Adds indentation to the XML tree.
    """
    i = "\n" + level*"  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
        for elem in elem:
            indent(elem, level+1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  Software Components __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def CreateSwcs():
    global swc_type
    # Retrieve the value from the swc_info dictionary at key 'B2'
    swc_type = swc_info['B2'].value
    
    # Define a switcher dictionary mapping component types to their corresponding functions
    switcher = {
        'ApplicationSwComponentType': my_application_function,  # Change the function name here
        'ComplexDeviceDriverSwComponentType': my_complex_device_driver_function,  # Change the function name here
        'EcuAbstractionSwComponentType': my_ecu_abstraction_function,  # Change the function name here
        'NvBlockSwComponentType': my_nv_block_function,  # Change the function name here
        'ParameterSwComponentType': my_parameter_function,  # Change the function name here
        'SensorActuatorSwComponentType': my_sensor_actuator_function,  # Change the function name here
        'ServiceProxySwComponentType': my_service_proxy_function,  # Change the function name here
        'ServiceSwComponentType': my_service_function  # Change the function name here
    }
    
    # Get the function from the switcher dictionary, defaulting to my_application_function
    func = switcher.get(swc_type, my_application_function)  # Change the function name here
    
    # Call the function
    func()

def my_application_function():

# ARXML structure
#   Appl SWC
#       Short Name
#       Ports
#           R port
#               short name
#               required interface
#           P port
#               short name
#               provided interface
#       IB
#           uuid
#           short name
#           constant memo
#           DTMS
#           Statis memo
#           AR type PIM
#           Events
#           Ex IRV
#           handleTerminationAndRestart
#           Imp IRV
#           per instance param
#           runnables
#           shared param
#           SupportsMultipleInstantiation
    

    # Declare global variables to be used within this function
    global currentfolder, CurrentSWC_shortname, CurrentInternalBehaviors
    
    currentfolder = 'ApplSWC' #other folders are CddSWC, CompSWC, EcuAbSWC, NvDataSWC, PrmSWC, SnsrActSWC, SrvcPrxySWC, SrvcSWC
    # Create a new application software component using the value from swc_info at key 'C2'
    CurrentSWC_shortname = (swc_info['C2'].value) 

    ApplSWC_folder_elements = arxml_structure.get_variable('ApplSWC_folder_elements')

    arelements_def.ApplicationSwComponentType(ApplSWC_folder_elements,CurrentSWC_shortname)

    Createports()

    # Add a new internal behavior to the current software component using the value from swc_info at key 'E2'
    CurrentInternalBehaviors = swc_info['E2'].value

    arelements_def.internal_behaviors(CurrentInternalBehaviors, swc_type)

    # Read columns B to H from ib_data to get runnable details

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy = excel_reader.read_columns(ib_data, 'B', 'H')

    if 'ConstantMemory' in IBVariableType:

        arelements_def.ConstantMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ConstantMemory':
                arelements_def.ConstantMemory_PDP(b, c, d, f, g)   
        
    else:
        print("ConstantMemorys are not present for this component")

    def createDTMS():
        # Read columns B & M from adt_primitive,  and columns B & H from adt_composite to get DTMS details
        adtp, cp, dp, ep, fp, gp, hp, ip, jp, kp, lp, idtp = excel_reader.read_columns(adt_primitive, 'B', 'M')

        adtc, cc, dc, ec, fc, gc, hc, idtc = excel_reader.read_columns(adt_composite, 'B', 'I')
        
        DataTypemappingSets_folder_elements = arxml_structure.get_variable('DataTypemappingSets_folder_elements')
        
        arelements_def.DataTypeMappingSet(DataTypemappingSets_folder_elements,CurrentSWC_shortname)

        for a,b in zip(adtp,idtp):
            arelements_def.data_type_map(a,b)

        for a,b in zip(adtc,idtc):
            arelements_def.data_type_map(a,b)

    createDTMS()

    arelements_def.DataTYPEMAPPINGREFS()
    arelements_def.DataTYPEMAPPINGREF(CurrentSWC_shortname)

    if 'StaticMemory' in IBVariableType:

        arelements_def.StaticMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'StaticMemory':
                arelements_def.StaticMemory_VDP(b, c, d, f, g)   
        
    else:
        print("StaticMemorys are not present for this component")

    if 'ArTypedPerInstanceMemory' in IBVariableType:

        arelements_def.ArTypedPerInstanceMemory()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ArTypedPerInstanceMemory':
                arelements_def.ArTypedPerInstanceMemory_VDP(b, c, d, f, g)   
        
    else:
        print("ArTypedPerInstanceMemorys are not present for this component")

    #createRTEEvents()

    rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo = excel_reader.read_columns(swc_info, 'H', 'M')

    # print(rnblname)

    arelements_def.RTE_Event()

    processed_types = set()

    for a,b,c,d,e,f in zip(rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo):

        if a in processed_types:
            continue
        processed_types.add(a)  
 
        # print(f"Processing: {a}, Category: {a}")

        # Check the type of RTE event and call the corresponding function
        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event
            arelements_def.AsynchronousServerCallReturnsEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'InitEvent':
            # Handle initialization event
            arelements_def.InitEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'BackgroundEvent':
            # Handle background event
            arelements_def.BackgroundEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.TimingEvent(d, a, currentfolder, CurrentSWC_shortname, f) #g is periodictime
        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information
            arelements_def.DataReceiveErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information
            arelements_def.DataSendCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information
            arelements_def.DataWriteCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information
            arelements_def.ExternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, trigger
        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information
            arelements_def.InternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information
            arelements_def.ModeSwitchedAckEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, modegroup
        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information
            arelements_def.OperationInvokedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, If_name, operation
        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information
            arelements_def.SwcModeManagerErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information
            arelements_def.DataReceivedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information
            arelements_def.SwcModeSwitchEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, modegroup, mode
        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information
            arelements_def.TransformerHardErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")


    if 'ExplicitInterRunnableVariable' in IBVariableType:

        arelements_def.ExplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ExplicitInterRunnableVariable':
                arelements_def.ExplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ExplicitInterRunnableVariable are not present for this component")

    # Set the handle Termination And Restart based on the value from swc_info at key 'F2'
    # CurrentInternalBehaviors.handleTerminationAndRestart = swc_info['F2'].value

    handleTerminationAndRestart= swc_info['F2'].value
    arelements_def.handle_termination_and_restart(handleTerminationAndRestart)

    if 'ImplicitInterRunnableVariables' in IBVariableType:

        arelements_def.ImplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ImplicitInterRunnableVariables':
                arelements_def.ImplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ImplicitInterRunnableVariables are not present for this component")

    if 'PerInstanceParameter' in IBVariableType:

        arelements_def.PerInstanceParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'PerInstanceParameter':
                arelements_def.PerInstanceParameter_PDP(b, c, d, f, g)   
        
    else:
        print("PerInstanceParameter are not present for this component")

    #create runnable here

    arelements_def.create_Runnable()

    processed_types = set()

    for a,b,c,d,e in zip(rnblname, rs, cic, rteeventname, rteeventtype):

        if b is None or (isinstance(b, str) and not b.strip()):
            b = a
        else :
            pass

        if a in processed_types:
            continue
        processed_types.add(a) 

        # Check the type of RTE event and call the corresponding function

        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event

            arelements_def.Runnable_ASCRE(a,currentfolder, CurrentSWC_shortname) #rport, If_name, operation

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InitEvent':
            # Handle initialization event

            arelements_def.Runnable_Init(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'BackgroundEvent':
            # Handle background event

            arelements_def.Runnable_BE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.Runnable_TE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information

            arelements_def.Runnable_DREE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information

            arelements_def.Runnable_DSCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information

            arelements_def.Runnable_DWCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information

            arelements_def.Runnable_ETOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information

            arelements_def.Runnable_ITOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information

            arelements_def.Runnable_MSAE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, modegroup,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information

            arelements_def.Runnable_OIE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information

            arelements_def.Runnable_SMMEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information

            arelements_def.Runnable_DRE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information

            arelements_def.Runnable_SMSE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information

            arelements_def.Runnable_THEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass
        
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")

    if 'SharedParameter' in IBVariableType:

        arelements_def.SharedParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'SharedParameter':
                arelements_def.SharedParameter_PDP(b, c, d, f, g)   
        
    else:
        print("SharedParameter are not present for this component")

    # Set the support for multiple instantiation based on the value from swc_info at key 'G2'

    SupportsMultipleInstantiation = swc_info['G2'].value

    arelements_def.supports_multiple_instantiation(SupportsMultipleInstantiation)

def my_complex_device_driver_function():

    # Declare global variables to be used within this function
    global currentfolder, CurrentSWC_shortname, CurrentInternalBehaviors
    
    currentfolder = 'CddSWC' #other folders are CddSWC, CompSWC, EcuAbSWC, NvDataSWC, PrmSWC, SnsrActSWC, SrvcPrxySWC, SrvcSWC
    # Create a new application software component using the value from swc_info at key 'C2'
    CurrentSWC_shortname = (swc_info['C2'].value) 

    CddSWC_folder_elements = arxml_structure.get_variable('CddSWC_folder_elements')

    arelements_def.ComplexDeviceDriverSwComponentType(CddSWC_folder_elements,CurrentSWC_shortname)

    Createports()

    # Add a new internal behavior to the current software component using the value from swc_info at key 'E2'
    CurrentInternalBehaviors = swc_info['E2'].value

    arelements_def.internal_behaviors(CurrentInternalBehaviors, swc_type)

    # Read columns B to H from ib_data to get runnable details

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy = excel_reader.read_columns(ib_data, 'B', 'H')

    if 'ConstantMemory' in IBVariableType:

        arelements_def.ConstantMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ConstantMemory':
                arelements_def.ConstantMemory_PDP(b, c, d, f, g)   
        
    else:
        print("ConstantMemorys are not present for this component")

    def createDTMS():
        # Read columns B & M from adt_primitive,  and columns B & H from adt_composite to get DTMS details
        adtp, cp, dp, ep, fp, gp, hp, ip, jp, kp, lp, idtp = excel_reader.read_columns(adt_primitive, 'B', 'M')

        adtc, cc, dc, ec, fc, gc, hc, idtc = excel_reader.read_columns(adt_composite, 'B', 'I')
        
        DataTypemappingSets_folder_elements = arxml_structure.get_variable('DataTypemappingSets_folder_elements')
        
        arelements_def.DataTypeMappingSet(DataTypemappingSets_folder_elements,CurrentSWC_shortname)

        for a,b in zip(adtp,idtp):
            arelements_def.data_type_map(a,b)

        for a,b in zip(adtc,idtc):
            arelements_def.data_type_map(a,b)

    createDTMS()

    arelements_def.DataTYPEMAPPINGREFS()
    arelements_def.DataTYPEMAPPINGREF(CurrentSWC_shortname)

    if 'StaticMemory' in IBVariableType:

        arelements_def.StaticMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'StaticMemory':
                arelements_def.StaticMemory_VDP(b, c, d, f, g)   
        
    else:
        print("StaticMemorys are not present for this component")

    if 'ArTypedPerInstanceMemory' in IBVariableType:

        arelements_def.ArTypedPerInstanceMemory()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ArTypedPerInstanceMemory':
                arelements_def.ArTypedPerInstanceMemory_VDP(b, c, d, f, g)   
        
    else:
        print("ArTypedPerInstanceMemorys are not present for this component")

    #createRTEEvents()

    rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo = excel_reader.read_columns(swc_info, 'H', 'M')

    # print(rnblname)

    arelements_def.RTE_Event()

    processed_types = set()

    for a,b,c,d,e,f in zip(rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo):

        if a in processed_types:
            continue
        processed_types.add(a)  
 
        # print(f"Processing: {a}, Category: {a}")

        # Check the type of RTE event and call the corresponding function
        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event
            arelements_def.AsynchronousServerCallReturnsEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'InitEvent':
            # Handle initialization event
            arelements_def.InitEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'BackgroundEvent':
            # Handle background event
            arelements_def.BackgroundEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.TimingEvent(d, a, currentfolder, CurrentSWC_shortname, f) #g is periodictime
        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information
            arelements_def.DataReceiveErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information
            arelements_def.DataSendCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information
            arelements_def.DataWriteCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information
            arelements_def.ExternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, trigger
        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information
            arelements_def.InternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information
            arelements_def.ModeSwitchedAckEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, modegroup
        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information
            arelements_def.OperationInvokedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, If_name, operation
        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information
            arelements_def.SwcModeManagerErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information
            arelements_def.DataReceivedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information
            arelements_def.SwcModeSwitchEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, modegroup, mode
        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information
            arelements_def.TransformerHardErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")


    if 'ExplicitInterRunnableVariable' in IBVariableType:

        arelements_def.ExplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ExplicitInterRunnableVariable':
                arelements_def.ExplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ExplicitInterRunnableVariable are not present for this component")

    # Set the handle Termination And Restart based on the value from swc_info at key 'F2'
    # CurrentInternalBehaviors.handleTerminationAndRestart = swc_info['F2'].value

    handleTerminationAndRestart= swc_info['F2'].value
    arelements_def.handle_termination_and_restart(handleTerminationAndRestart)

    if 'ImplicitInterRunnableVariables' in IBVariableType:

        arelements_def.ImplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ImplicitInterRunnableVariables':
                arelements_def.ImplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ImplicitInterRunnableVariables are not present for this component")

    if 'PerInstanceParameter' in IBVariableType:

        arelements_def.PerInstanceParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'PerInstanceParameter':
                arelements_def.PerInstanceParameter_PDP(b, c, d, f, g)   
        
    else:
        print("PerInstanceParameter are not present for this component")

    #create runnable here

    arelements_def.create_Runnable()

    processed_types = set()

    for a,b,c,d,e in zip(rnblname, rs, cic, rteeventname, rteeventtype):

        if b is None or (isinstance(b, str) and not b.strip()):
            b = a
        else :
            pass

        if a in processed_types:
            continue
        processed_types.add(a) 

        # Check the type of RTE event and call the corresponding function

        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event

            arelements_def.Runnable_ASCRE(a,currentfolder, CurrentSWC_shortname) #rport, If_name, operation

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InitEvent':
            # Handle initialization event

            arelements_def.Runnable_Init(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'BackgroundEvent':
            # Handle background event

            arelements_def.Runnable_BE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.Runnable_TE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information

            arelements_def.Runnable_DREE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information

            arelements_def.Runnable_DSCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information

            arelements_def.Runnable_DWCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information

            arelements_def.Runnable_ETOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information

            arelements_def.Runnable_ITOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information

            arelements_def.Runnable_MSAE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, modegroup,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information

            arelements_def.Runnable_OIE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information

            arelements_def.Runnable_SMMEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information

            arelements_def.Runnable_DRE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information

            arelements_def.Runnable_SMSE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information

            arelements_def.Runnable_THEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass
        
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")

    if 'SharedParameter' in IBVariableType:

        arelements_def.SharedParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'SharedParameter':
                arelements_def.SharedParameter_PDP(b, c, d, f, g)   
        
    else:
        print("SharedParameter are not present for this component")

    # Set the support for multiple instantiation based on the value from swc_info at key 'G2'

    SupportsMultipleInstantiation = swc_info['G2'].value

    arelements_def.supports_multiple_instantiation(SupportsMultipleInstantiation)

def my_ecu_abstraction_function():

    # Declare global variables to be used within this function
    global currentfolder, CurrentSWC_shortname, CurrentInternalBehaviors
    
    currentfolder = 'EcuAbSWC' #other folders are CddSWC, CompSWC, EcuAbSWC, NvDataSWC, PrmSWC, SnsrActSWC, SrvcPrxySWC, SrvcSWC
    # Create a new application software component using the value from swc_info at key 'C2'
    CurrentSWC_shortname = (swc_info['C2'].value) 

    EcuAbSWC_folder_elements = arxml_structure.get_variable('EcuAbSWC_folder_elements')

    arelements_def.EcuAbstractionSwComponentType(EcuAbSWC_folder_elements,CurrentSWC_shortname)

    Createports()

    # Add a new internal behavior to the current software component using the value from swc_info at key 'E2'
    CurrentInternalBehaviors = swc_info['E2'].value

    arelements_def.internal_behaviors(CurrentInternalBehaviors, swc_type)

    # Read columns B to H from ib_data to get runnable details

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy = excel_reader.read_columns(ib_data, 'B', 'H')

    if 'ConstantMemory' in IBVariableType:

        arelements_def.ConstantMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ConstantMemory':
                arelements_def.ConstantMemory_PDP(b, c, d, f, g)   
        
    else:
        print("ConstantMemorys are not present for this component")

    def createDTMS():
        # Read columns B & M from adt_primitive,  and columns B & H from adt_composite to get DTMS details
        adtp, cp, dp, ep, fp, gp, hp, ip, jp, kp, lp, idtp = excel_reader.read_columns(adt_primitive, 'B', 'M')

        adtc, cc, dc, ec, fc, gc, hc, idtc = excel_reader.read_columns(adt_composite, 'B', 'I')
        
        DataTypemappingSets_folder_elements = arxml_structure.get_variable('DataTypemappingSets_folder_elements')
        
        arelements_def.DataTypeMappingSet(DataTypemappingSets_folder_elements,CurrentSWC_shortname)

        for a,b in zip(adtp,idtp):
            arelements_def.data_type_map(a,b)

        for a,b in zip(adtc,idtc):
            arelements_def.data_type_map(a,b)

    createDTMS()

    arelements_def.DataTYPEMAPPINGREFS()
    arelements_def.DataTYPEMAPPINGREF(CurrentSWC_shortname)

    if 'StaticMemory' in IBVariableType:

        arelements_def.StaticMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'StaticMemory':
                arelements_def.StaticMemory_VDP(b, c, d, f, g)   
        
    else:
        print("StaticMemorys are not present for this component")

    if 'ArTypedPerInstanceMemory' in IBVariableType:

        arelements_def.ArTypedPerInstanceMemory()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ArTypedPerInstanceMemory':
                arelements_def.ArTypedPerInstanceMemory_VDP(b, c, d, f, g)   
        
    else:
        print("ArTypedPerInstanceMemorys are not present for this component")

    #createRTEEvents()

    rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo = excel_reader.read_columns(swc_info, 'H', 'M')

    # print(rnblname)

    arelements_def.RTE_Event()

    processed_types = set()

    for a,b,c,d,e,f in zip(rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo):

        if a in processed_types:
            continue
        processed_types.add(a)  
 
        # print(f"Processing: {a}, Category: {a}")

        # Check the type of RTE event and call the corresponding function
        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event
            arelements_def.AsynchronousServerCallReturnsEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'InitEvent':
            # Handle initialization event
            arelements_def.InitEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'BackgroundEvent':
            # Handle background event
            arelements_def.BackgroundEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.TimingEvent(d, a, currentfolder, CurrentSWC_shortname, f) #g is periodictime
        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information
            arelements_def.DataReceiveErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information
            arelements_def.DataSendCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information
            arelements_def.DataWriteCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information
            arelements_def.ExternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, trigger
        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information
            arelements_def.InternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information
            arelements_def.ModeSwitchedAckEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, modegroup
        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information
            arelements_def.OperationInvokedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, If_name, operation
        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information
            arelements_def.SwcModeManagerErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information
            arelements_def.DataReceivedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information
            arelements_def.SwcModeSwitchEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, modegroup, mode
        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information
            arelements_def.TransformerHardErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")


    if 'ExplicitInterRunnableVariable' in IBVariableType:

        arelements_def.ExplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ExplicitInterRunnableVariable':
                arelements_def.ExplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ExplicitInterRunnableVariable are not present for this component")

    # Set the handle Termination And Restart based on the value from swc_info at key 'F2'
    # CurrentInternalBehaviors.handleTerminationAndRestart = swc_info['F2'].value

    handleTerminationAndRestart= swc_info['F2'].value
    arelements_def.handle_termination_and_restart(handleTerminationAndRestart)

    if 'ImplicitInterRunnableVariables' in IBVariableType:

        arelements_def.ImplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ImplicitInterRunnableVariables':
                arelements_def.ImplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ImplicitInterRunnableVariables are not present for this component")

    if 'PerInstanceParameter' in IBVariableType:

        arelements_def.PerInstanceParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'PerInstanceParameter':
                arelements_def.PerInstanceParameter_PDP(b, c, d, f, g)   
        
    else:
        print("PerInstanceParameter are not present for this component")

    #create runnable here

    arelements_def.create_Runnable()

    processed_types = set()

    for a,b,c,d,e in zip(rnblname, rs, cic, rteeventname, rteeventtype):

        if b is None or (isinstance(b, str) and not b.strip()):
            b = a
        else :
            pass

        if a in processed_types:
            continue
        processed_types.add(a) 

        # Check the type of RTE event and call the corresponding function

        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event

            arelements_def.Runnable_ASCRE(a,currentfolder, CurrentSWC_shortname) #rport, If_name, operation

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InitEvent':
            # Handle initialization event

            arelements_def.Runnable_Init(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'BackgroundEvent':
            # Handle background event

            arelements_def.Runnable_BE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.Runnable_TE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information

            arelements_def.Runnable_DREE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information

            arelements_def.Runnable_DSCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information

            arelements_def.Runnable_DWCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information

            arelements_def.Runnable_ETOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information

            arelements_def.Runnable_ITOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information

            arelements_def.Runnable_MSAE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, modegroup,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information

            arelements_def.Runnable_OIE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information

            arelements_def.Runnable_SMMEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information

            arelements_def.Runnable_DRE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information

            arelements_def.Runnable_SMSE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information

            arelements_def.Runnable_THEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass
        
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")

    if 'SharedParameter' in IBVariableType:

        arelements_def.SharedParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'SharedParameter':
                arelements_def.SharedParameter_PDP(b, c, d, f, g)   
        
    else:
        print("SharedParameter are not present for this component")

    # Set the support for multiple instantiation based on the value from swc_info at key 'G2'

    SupportsMultipleInstantiation = swc_info['G2'].value

    arelements_def.supports_multiple_instantiation(SupportsMultipleInstantiation)

def my_nv_block_function():
    print('my_nv_block_function : will create later')

def my_parameter_function():
    print('my_parameter_function : will create later')

def my_sensor_actuator_function():

    # Declare global variables to be used within this function
    global currentfolder, CurrentSWC_shortname, CurrentInternalBehaviors
    
    currentfolder = 'SnsrActSWC' #other folders are CddSWC, CompSWC, EcuAbSWC, NvDataSWC, PrmSWC, SnsrActSWC, SrvcPrxySWC, SrvcSWC
    # Create a new application software component using the value from swc_info at key 'C2'
    CurrentSWC_shortname = (swc_info['C2'].value) 

    SnsrActSWC_folder_elements = arxml_structure.get_variable('SnsrActSWC_folder_elements')

    arelements_def.SensorActuatorSwComponentType(SnsrActSWC_folder_elements,CurrentSWC_shortname)

    Createports()

    # Add a new internal behavior to the current software component using the value from swc_info at key 'E2'
    CurrentInternalBehaviors = swc_info['E2'].value

    arelements_def.internal_behaviors(CurrentInternalBehaviors, swc_type)

    # Read columns B to H from ib_data to get runnable details

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy = excel_reader.read_columns(ib_data, 'B', 'H')

    if 'ConstantMemory' in IBVariableType:

        arelements_def.ConstantMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ConstantMemory':
                arelements_def.ConstantMemory_PDP(b, c, d, f, g)   
        
    else:
        print("ConstantMemorys are not present for this component")

    def createDTMS():
        # Read columns B & M from adt_primitive,  and columns B & H from adt_composite to get DTMS details
        adtp, cp, dp, ep, fp, gp, hp, ip, jp, kp, lp, idtp = excel_reader.read_columns(adt_primitive, 'B', 'M')

        adtc, cc, dc, ec, fc, gc, hc, idtc = excel_reader.read_columns(adt_composite, 'B', 'I')
        
        DataTypemappingSets_folder_elements = arxml_structure.get_variable('DataTypemappingSets_folder_elements')
        
        arelements_def.DataTypeMappingSet(DataTypemappingSets_folder_elements,CurrentSWC_shortname)

        for a,b in zip(adtp,idtp):
            arelements_def.data_type_map(a,b)

        for a,b in zip(adtc,idtc):
            arelements_def.data_type_map(a,b)

    createDTMS()

    arelements_def.DataTYPEMAPPINGREFS()
    arelements_def.DataTYPEMAPPINGREF(CurrentSWC_shortname)

    if 'StaticMemory' in IBVariableType:

        arelements_def.StaticMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'StaticMemory':
                arelements_def.StaticMemory_VDP(b, c, d, f, g)   
        
    else:
        print("StaticMemorys are not present for this component")

    if 'ArTypedPerInstanceMemory' in IBVariableType:

        arelements_def.ArTypedPerInstanceMemory()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ArTypedPerInstanceMemory':
                arelements_def.ArTypedPerInstanceMemory_VDP(b, c, d, f, g)   
        
    else:
        print("ArTypedPerInstanceMemorys are not present for this component")

    #createRTEEvents()

    rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo = excel_reader.read_columns(swc_info, 'H', 'M')

    # print(rnblname)

    arelements_def.RTE_Event()

    processed_types = set()

    for a,b,c,d,e,f in zip(rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo):

        if a in processed_types:
            continue
        processed_types.add(a)  
 
        # print(f"Processing: {a}, Category: {a}")

        # Check the type of RTE event and call the corresponding function
        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event
            arelements_def.AsynchronousServerCallReturnsEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'InitEvent':
            # Handle initialization event
            arelements_def.InitEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'BackgroundEvent':
            # Handle background event
            arelements_def.BackgroundEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.TimingEvent(d, a, currentfolder, CurrentSWC_shortname, f) #g is periodictime
        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information
            arelements_def.DataReceiveErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information
            arelements_def.DataSendCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information
            arelements_def.DataWriteCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information
            arelements_def.ExternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, trigger
        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information
            arelements_def.InternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information
            arelements_def.ModeSwitchedAckEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, modegroup
        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information
            arelements_def.OperationInvokedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, If_name, operation
        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information
            arelements_def.SwcModeManagerErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information
            arelements_def.DataReceivedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information
            arelements_def.SwcModeSwitchEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, modegroup, mode
        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information
            arelements_def.TransformerHardErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")


    if 'ExplicitInterRunnableVariable' in IBVariableType:

        arelements_def.ExplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ExplicitInterRunnableVariable':
                arelements_def.ExplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ExplicitInterRunnableVariable are not present for this component")

    # Set the handle Termination And Restart based on the value from swc_info at key 'F2'
    # CurrentInternalBehaviors.handleTerminationAndRestart = swc_info['F2'].value

    handleTerminationAndRestart= swc_info['F2'].value
    arelements_def.handle_termination_and_restart(handleTerminationAndRestart)

    if 'ImplicitInterRunnableVariables' in IBVariableType:

        arelements_def.ImplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ImplicitInterRunnableVariables':
                arelements_def.ImplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ImplicitInterRunnableVariables are not present for this component")

    if 'PerInstanceParameter' in IBVariableType:

        arelements_def.PerInstanceParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'PerInstanceParameter':
                arelements_def.PerInstanceParameter_PDP(b, c, d, f, g)   
        
    else:
        print("PerInstanceParameter are not present for this component")

    #create runnable here

    arelements_def.create_Runnable()

    processed_types = set()

    for a,b,c,d,e in zip(rnblname, rs, cic, rteeventname, rteeventtype):

        if b is None or (isinstance(b, str) and not b.strip()):
            b = a
        else :
            pass

        if a in processed_types:
            continue
        processed_types.add(a) 

        # Check the type of RTE event and call the corresponding function

        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event

            arelements_def.Runnable_ASCRE(a,currentfolder, CurrentSWC_shortname) #rport, If_name, operation

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InitEvent':
            # Handle initialization event

            arelements_def.Runnable_Init(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'BackgroundEvent':
            # Handle background event

            arelements_def.Runnable_BE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.Runnable_TE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information

            arelements_def.Runnable_DREE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information

            arelements_def.Runnable_DSCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information

            arelements_def.Runnable_DWCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information

            arelements_def.Runnable_ETOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information

            arelements_def.Runnable_ITOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information

            arelements_def.Runnable_MSAE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, modegroup,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information

            arelements_def.Runnable_OIE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information

            arelements_def.Runnable_SMMEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information

            arelements_def.Runnable_DRE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information

            arelements_def.Runnable_SMSE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information

            arelements_def.Runnable_THEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass
        
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")

    if 'SharedParameter' in IBVariableType:

        arelements_def.SharedParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'SharedParameter':
                arelements_def.SharedParameter_PDP(b, c, d, f, g)   
        
    else:
        print("SharedParameter are not present for this component")

    # Set the support for multiple instantiation based on the value from swc_info at key 'G2'

    SupportsMultipleInstantiation = swc_info['G2'].value

    arelements_def.supports_multiple_instantiation(SupportsMultipleInstantiation)

def my_service_proxy_function():

    # Declare global variables to be used within this function
    global currentfolder, CurrentSWC_shortname, CurrentInternalBehaviors
    
    currentfolder = 'SrvcPrxySWC' #other folders are CddSWC, CompSWC, EcuAbSWC, NvDataSWC, PrmSWC, SnsrActSWC, SrvcPrxySWC, SrvcSWC
    # Create a new application software component using the value from swc_info at key 'C2'
    CurrentSWC_shortname = (swc_info['C2'].value) 

    SrvcPrxySWC_folder_elements = arxml_structure.get_variable('SrvcPrxySWC_folder_elements')

    arelements_def.ServiceProxySwComponentType(SrvcPrxySWC_folder_elements,CurrentSWC_shortname)

    Createports()

    # Add a new internal behavior to the current software component using the value from swc_info at key 'E2'
    CurrentInternalBehaviors = swc_info['E2'].value

    arelements_def.internal_behaviors(CurrentInternalBehaviors, swc_type)

    # Read columns B to H from ib_data to get runnable details

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy = excel_reader.read_columns(ib_data, 'B', 'H')

    if 'ConstantMemory' in IBVariableType:

        arelements_def.ConstantMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ConstantMemory':
                arelements_def.ConstantMemory_PDP(b, c, d, f, g)   
        
    else:
        print("ConstantMemorys are not present for this component")

    def createDTMS():
        # Read columns B & M from adt_primitive,  and columns B & H from adt_composite to get DTMS details
        adtp, cp, dp, ep, fp, gp, hp, ip, jp, kp, lp, idtp = excel_reader.read_columns(adt_primitive, 'B', 'M')

        adtc, cc, dc, ec, fc, gc, hc, idtc = excel_reader.read_columns(adt_composite, 'B', 'I')
        
        DataTypemappingSets_folder_elements = arxml_structure.get_variable('DataTypemappingSets_folder_elements')
        
        arelements_def.DataTypeMappingSet(DataTypemappingSets_folder_elements,CurrentSWC_shortname)

        for a,b in zip(adtp,idtp):
            arelements_def.data_type_map(a,b)

        for a,b in zip(adtc,idtc):
            arelements_def.data_type_map(a,b)

    createDTMS()

    arelements_def.DataTYPEMAPPINGREFS()
    arelements_def.DataTYPEMAPPINGREF(CurrentSWC_shortname)

    if 'StaticMemory' in IBVariableType:

        arelements_def.StaticMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'StaticMemory':
                arelements_def.StaticMemory_VDP(b, c, d, f, g)   
        
    else:
        print("StaticMemorys are not present for this component")

    if 'ArTypedPerInstanceMemory' in IBVariableType:

        arelements_def.ArTypedPerInstanceMemory()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ArTypedPerInstanceMemory':
                arelements_def.ArTypedPerInstanceMemory_VDP(b, c, d, f, g)   
        
    else:
        print("ArTypedPerInstanceMemorys are not present for this component")

    #createRTEEvents()

    rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo = excel_reader.read_columns(swc_info, 'H', 'M')

    # print(rnblname)

    arelements_def.RTE_Event()

    processed_types = set()

    for a,b,c,d,e,f in zip(rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo):

        if a in processed_types:
            continue
        processed_types.add(a)  
 
        # print(f"Processing: {a}, Category: {a}")

        # Check the type of RTE event and call the corresponding function
        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event
            arelements_def.AsynchronousServerCallReturnsEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'InitEvent':
            # Handle initialization event
            arelements_def.InitEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'BackgroundEvent':
            # Handle background event
            arelements_def.BackgroundEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.TimingEvent(d, a, currentfolder, CurrentSWC_shortname, f) #g is periodictime
        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information
            arelements_def.DataReceiveErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information
            arelements_def.DataSendCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information
            arelements_def.DataWriteCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information
            arelements_def.ExternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, trigger
        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information
            arelements_def.InternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information
            arelements_def.ModeSwitchedAckEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, modegroup
        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information
            arelements_def.OperationInvokedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, If_name, operation
        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information
            arelements_def.SwcModeManagerErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information
            arelements_def.DataReceivedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information
            arelements_def.SwcModeSwitchEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, modegroup, mode
        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information
            arelements_def.TransformerHardErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")


    if 'ExplicitInterRunnableVariable' in IBVariableType:

        arelements_def.ExplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ExplicitInterRunnableVariable':
                arelements_def.ExplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ExplicitInterRunnableVariable are not present for this component")

    # Set the handle Termination And Restart based on the value from swc_info at key 'F2'
    # CurrentInternalBehaviors.handleTerminationAndRestart = swc_info['F2'].value

    handleTerminationAndRestart= swc_info['F2'].value
    arelements_def.handle_termination_and_restart(handleTerminationAndRestart)

    if 'ImplicitInterRunnableVariables' in IBVariableType:

        arelements_def.ImplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ImplicitInterRunnableVariables':
                arelements_def.ImplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ImplicitInterRunnableVariables are not present for this component")

    if 'PerInstanceParameter' in IBVariableType:

        arelements_def.PerInstanceParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'PerInstanceParameter':
                arelements_def.PerInstanceParameter_PDP(b, c, d, f, g)   
        
    else:
        print("PerInstanceParameter are not present for this component")

    #create runnable here

    arelements_def.create_Runnable()

    processed_types = set()

    for a,b,c,d,e in zip(rnblname, rs, cic, rteeventname, rteeventtype):

        if b is None or (isinstance(b, str) and not b.strip()):
            b = a
        else :
            pass

        if a in processed_types:
            continue
        processed_types.add(a) 

        # Check the type of RTE event and call the corresponding function

        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event

            arelements_def.Runnable_ASCRE(a,currentfolder, CurrentSWC_shortname) #rport, If_name, operation

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InitEvent':
            # Handle initialization event

            arelements_def.Runnable_Init(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'BackgroundEvent':
            # Handle background event

            arelements_def.Runnable_BE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.Runnable_TE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information

            arelements_def.Runnable_DREE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information

            arelements_def.Runnable_DSCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information

            arelements_def.Runnable_DWCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information

            arelements_def.Runnable_ETOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information

            arelements_def.Runnable_ITOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information

            arelements_def.Runnable_MSAE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, modegroup,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information

            arelements_def.Runnable_OIE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information

            arelements_def.Runnable_SMMEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information

            arelements_def.Runnable_DRE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information

            arelements_def.Runnable_SMSE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information

            arelements_def.Runnable_THEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass
        
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")

    if 'SharedParameter' in IBVariableType:

        arelements_def.SharedParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'SharedParameter':
                arelements_def.SharedParameter_PDP(b, c, d, f, g)   
        
    else:
        print("SharedParameter are not present for this component")

    # Set the support for multiple instantiation based on the value from swc_info at key 'G2'

    SupportsMultipleInstantiation = swc_info['G2'].value

    arelements_def.supports_multiple_instantiation(SupportsMultipleInstantiation)

def my_service_function():

    # Declare global variables to be used within this function
    global currentfolder, CurrentSWC_shortname, CurrentInternalBehaviors
    
    currentfolder = 'SrvcSWC' #other folders are CddSWC, CompSWC, EcuAbSWC, NvDataSWC, PrmSWC, SnsrActSWC, SrvcPrxySWC, SrvcSWC
    # Create a new application software component using the value from swc_info at key 'C2'
    CurrentSWC_shortname = (swc_info['C2'].value) 

    SrvcSWC_folder_elements = arxml_structure.get_variable('SrvcSWC_folder_elements')

    arelements_def.ServiceSwComponentType(SrvcSWC_folder_elements,CurrentSWC_shortname)

    Createports()

    # Add a new internal behavior to the current software component using the value from swc_info at key 'E2'
    CurrentInternalBehaviors = swc_info['E2'].value

    arelements_def.internal_behaviors(CurrentInternalBehaviors, swc_type)

    # Read columns B to H from ib_data to get runnable details

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy = excel_reader.read_columns(ib_data, 'B', 'H')

    if 'ConstantMemory' in IBVariableType:

        arelements_def.ConstantMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ConstantMemory':
                arelements_def.ConstantMemory_PDP(b, c, d, f, g)   
        
    else:
        print("ConstantMemorys are not present for this component")

    def createDTMS():
        # Read columns B & M from adt_primitive,  and columns B & H from adt_composite to get DTMS details
        adtp, cp, dp, ep, fp, gp, hp, ip, jp, kp, lp, idtp = excel_reader.read_columns(adt_primitive, 'B', 'M')

        adtc, cc, dc, ec, fc, gc, hc, idtc = excel_reader.read_columns(adt_composite, 'B', 'I')
        
        DataTypemappingSets_folder_elements = arxml_structure.get_variable('DataTypemappingSets_folder_elements')
        
        arelements_def.DataTypeMappingSet(DataTypemappingSets_folder_elements,CurrentSWC_shortname)

        for a,b in zip(adtp,idtp):
            arelements_def.data_type_map(a,b)

        for a,b in zip(adtc,idtc):
            arelements_def.data_type_map(a,b)

    createDTMS()

    arelements_def.DataTYPEMAPPINGREFS()
    arelements_def.DataTYPEMAPPINGREF(CurrentSWC_shortname)

    if 'StaticMemory' in IBVariableType:

        arelements_def.StaticMemory()
      
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'StaticMemory':
                arelements_def.StaticMemory_VDP(b, c, d, f, g)   
        
    else:
        print("StaticMemorys are not present for this component")

    if 'ArTypedPerInstanceMemory' in IBVariableType:

        arelements_def.ArTypedPerInstanceMemory()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ArTypedPerInstanceMemory':
                arelements_def.ArTypedPerInstanceMemory_VDP(b, c, d, f, g)   
        
    else:
        print("ArTypedPerInstanceMemorys are not present for this component")

    #createRTEEvents()

    rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo = excel_reader.read_columns(swc_info, 'H', 'M')

    # print(rnblname)

    arelements_def.RTE_Event()

    processed_types = set()

    for a,b,c,d,e,f in zip(rnblname, rs, cic, rteeventname, rteeventtype, rteeventinfo):

        if a in processed_types:
            continue
        processed_types.add(a)  
 
        # print(f"Processing: {a}, Category: {a}")

        # Check the type of RTE event and call the corresponding function
        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event
            arelements_def.AsynchronousServerCallReturnsEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'InitEvent':
            # Handle initialization event
            arelements_def.InitEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'BackgroundEvent':
            # Handle background event
            arelements_def.BackgroundEvent(d, a, currentfolder, CurrentSWC_shortname)
        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.TimingEvent(d, a, currentfolder, CurrentSWC_shortname, f) #g is periodictime
        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information
            arelements_def.DataReceiveErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information
            arelements_def.DataSendCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information
            arelements_def.DataWriteCompletedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, DE
        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information
            arelements_def.ExternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, trigger
        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information
            arelements_def.InternalTriggerOccurredEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information
            arelements_def.ModeSwitchedAckEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, modegroup
        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information
            arelements_def.OperationInvokedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #pport, If_name, operation
        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information
            arelements_def.SwcModeManagerErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information
            arelements_def.DataReceivedEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, DE
        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information
            arelements_def.SwcModeSwitchEvent(d, a, currentfolder, CurrentSWC_shortname, f) #rport, If_name, modegroup, mode
        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information
            arelements_def.TransformerHardErrorEvent(d, a, currentfolder, CurrentSWC_shortname, f)
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")


    if 'ExplicitInterRunnableVariable' in IBVariableType:

        arelements_def.ExplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ExplicitInterRunnableVariable':
                arelements_def.ExplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ExplicitInterRunnableVariable are not present for this component")

    # Set the handle Termination And Restart based on the value from swc_info at key 'F2'
    # CurrentInternalBehaviors.handleTerminationAndRestart = swc_info['F2'].value

    handleTerminationAndRestart= swc_info['F2'].value
    arelements_def.handle_termination_and_restart(handleTerminationAndRestart)

    if 'ImplicitInterRunnableVariables' in IBVariableType:

        arelements_def.ImplicitInterRunnableVariable()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'ImplicitInterRunnableVariables':
                arelements_def.ImplicitInterRunnableVariable_VDP(b, c, d, f, g)   
        
    else:
        print("ImplicitInterRunnableVariables are not present for this component")

    if 'PerInstanceParameter' in IBVariableType:

        arelements_def.PerInstanceParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'PerInstanceParameter':
                arelements_def.PerInstanceParameter_PDP(b, c, d, f, g)   
        
    else:
        print("PerInstanceParameter are not present for this component")

    #create runnable here

    arelements_def.create_Runnable()

    processed_types = set()

    for a,b,c,d,e in zip(rnblname, rs, cic, rteeventname, rteeventtype):

        if b is None or (isinstance(b, str) and not b.strip()):
            b = a
        else :
            pass

        if a in processed_types:
            continue
        processed_types.add(a) 

        # Check the type of RTE event and call the corresponding function

        if e == 'AsynchronousServerCallReturnsEvent':
            # Handle asynchronous server call return event

            arelements_def.Runnable_ASCRE(a,currentfolder, CurrentSWC_shortname) #rport, If_name, operation

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InitEvent':
            # Handle initialization event

            arelements_def.Runnable_Init(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'BackgroundEvent':
            # Handle background event

            arelements_def.Runnable_BE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TimingEvent':
            # Handle timing event with additional information
            arelements_def.Runnable_TE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceiveErrorEvent':
            # Handle data receive error event with additional information

            arelements_def.Runnable_DREE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataSendCompletedEvent':
            # Handle data send completed event with additional information

            arelements_def.Runnable_DSCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataWriteCompletedEvent':
            # Handle data write completed event with additional information

            arelements_def.Runnable_DWCE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, DE,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ExternalTriggerOccurredEvent':
            # Handle external trigger occurred event with additional information

            arelements_def.Runnable_ETOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'InternalTriggerOccurredEvent':
            # Handle internal trigger occurred event with additional information

            arelements_def.Runnable_ITOE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'ModeSwitchedAckEvent':
            # Handle mode switched acknowledgment event with additional information

            arelements_def.Runnable_MSAE(a,currentfolder, CurrentSWC_shortname)#pport, If_name, modegroup,

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'OperationInvokedEvent':
            # Handle operation invoked event with additional information

            arelements_def.Runnable_OIE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeManagerErrorEvent':
            # Handle software component mode manager error event with additional information

            arelements_def.Runnable_SMMEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'DataReceivedEvent':
            # Handle data received event with additional information

            arelements_def.Runnable_DRE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'SwcModeSwitchEvent':
            # Handle software component mode switch event with additional information

            arelements_def.Runnable_SMSE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass

        elif e == 'TransformerHardErrorEvent':
            # Handle transformer hard error event with additional information

            arelements_def.Runnable_THEE(a)

            m = excel_reader.read_columns(ib_data, 'F', 'F') or []  # Ensure m is at least an empty list
            n = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure n is at least an empty list

            # Ensure m and n are lists (in case the function returns None or something unexpected)
            if not isinstance(m, list):
                m = list(m)  
            if not isinstance(n, list):
                n = list(n)

            # Flatten m and n if they contain nested lists
            m = [item for sublist in m for item in sublist] if any(isinstance(i, list) for i in m) else m
            n = [item for sublist in n for item in sublist] if any(isinstance(i, list) for i in n) else n

            if a in m or a in n :
                rnblaccess(a)            

            arelements_def.Rnblsymbol(b)

            if a in m :
                rnblaccess_WrittenIRV(a)
            else :
                pass
        
        else:
            # Print a message for unrecognized event types
            print(f"Unrecognized event type: {e}")

    if 'SharedParameter' in IBVariableType:

        arelements_def.SharedParameter()
    
        for a,b,c,d,e,f,g in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable, SwCalibrationAccess, SwImplementationPolicy):
            if a == 'SharedParameter':
                arelements_def.SharedParameter_PDP(b, c, d, f, g)   
        
    else:
        print("SharedParameter are not present for this component")

    # Set the support for multiple instantiation based on the value from swc_info at key 'G2'

    SupportsMultipleInstantiation = swc_info['G2'].value

    arelements_def.supports_multiple_instantiation(SupportsMultipleInstantiation)



#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  runnable data access __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# data read access
# data receive point by arguments
# data receive point by values
# data send points
# data write access
# mode switch access
# Parameter access
    # Constant memory
    # per instance parameter
    # port parameter
    # shared parameter
# Server call Points
# Read local variables
    # explicit
    # implicit
# -----------Symbol------------#
# written local variables
    # explicit
    # implicit

# ib_data f column , ports i columns

def rnblaccess(Currentrnbl):

    print("\n[DEBUG] Starting rnblaccess function")
    print(f"[DEBUG] Processing for Current Rnbl: {Currentrnbl}")
    print(f"[DEBUG] Current folder: {currentfolder}, Current SWC shortname: {CurrentSWC_shortname}")
    
    # Read argument_col (Column G) from Excel
    argument_col = excel_reader.read_columns(ports, 'G', 'G') or []
    
    # Debugging: Print raw argument data
    print(f"[DEBUG] Raw Argument Column Data: {argument_col}")
    
    # Flatten the list to remove any nested lists
    def flatten(lst):
        return [item for sublist in lst for item in sublist] if any(isinstance(i, list) for i in lst) else lst
    
    argument_col = flatten(argument_col)
    
    # Debugging: Print after flattening
    print(f"[DEBUG] Flattened Argument Column: {argument_col}")

    #interface check

    inf_col = excel_reader.read_columns(ports, 'D', 'D') or []  # Ensure argument_col is at least an empty list    

    # Debugging: Print raw argument data
    print(f"[DEBUG] Raw Interface Column Data: {inf_col}")

    inf_col = flatten(inf_col)
    
    # Debugging: Print after flattening
    print(f"[DEBUG] Flattened Interface Column: {inf_col}")

    #accessing runnable check : ports

    ports_Acc_Rnbl = excel_reader.read_columns(ports, 'I', 'I') or []  # Ensure argument_col is at least an empty list    

    # Debugging: Print raw argument data
    print(f"[DEBUG] Raw accessing runnable from ports sheet Data: {ports_Acc_Rnbl}")

    ports_Acc_Rnbl = flatten(ports_Acc_Rnbl)
    
    # Debugging: Print after flattening
    print(f"[DEBUG] Flattened accessing runnable from ports Column: {ports_Acc_Rnbl}")

    # Call main function if any predefined argument exists

    for a,b in zip(argument_col,  ports_Acc_Rnbl):

        if 'dra' in a and b == Currentrnbl :
            print(f"[DEBUG] Found 'dra', calling arelements_def.dra()")
            arelements_def.dra()
        elif 'drpa' in a and b == Currentrnbl:
            print(f"[DEBUG] Found 'drpa', calling arelements_def.drpa()")
            arelements_def.drpa()
        elif 'drpv' in a and b == Currentrnbl:
            print(f"[DEBUG] Found 'drpv', calling arelements_def.drpv()")
            arelements_def.drpv()
        elif 'dsp' in a and b == Currentrnbl:
            print(f"[DEBUG] Found 'dsp', calling arelements_def.dsp()")
            arelements_def.dsp()
        elif 'dwa' in a and b == Currentrnbl:
            print(f"[DEBUG] Found 'dwa', calling arelements_def.dwa()")
            arelements_def.dwa()
        else :
            pass

    global pa_already_triggered

    pa_already_triggered = None

    for a,b in zip(inf_col,  ports_Acc_Rnbl):

        if 'ModeSwitchInterface' in a and b == Currentrnbl:
            print(f"[DEBUG] Found 'ModeSwitchInterface', calling arelements_def.msp()")
            arelements_def.msp()        
             
        elif 'ParameterInterface' in a and b == Currentrnbl :
            print(f"[DEBUG] Found one of the 'ParameterInterface' 'ConstantMemory', 'PerInstanceParameter', 'SharedParameter', calling arelements_def.pa()")
            arelements_def.pa()
            pa_already_triggered = 1

        elif 'ClientServerInterface' in a and b == Currentrnbl:
            print(f"[DEBUG] Found 'ClientServerInterface', calling arelements_def.sscp()")
            arelements_def.sscp()

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable = excel_reader.read_columns(ib_data, 'B', 'F')

    for ait,br in zip(IBVariableType, AccessingRunnable):
        if pa_already_triggered != 1 :

            if ait in ['ConstantMemory', 'PerInstanceParameter', 'SharedParameter'] and br == Currentrnbl :
                arelements_def.pa()
                pa_already_triggered = 0
                break

    for irvt,br in zip(IBVariableType, AccessingRunnable):
        if irvt in ['ImplicitInterRunnableVariables', 'ExplicitInterRunnableVariable'] and br == Currentrnbl :#still read or write bifurcation is pending
            print(f"[DEBUG] Found 'ImplicitInterRunnableVariables', calling arelements_def.IRVRA()")        
            arelements_def.IRVRA()
            break

    for a,b,c,d,e in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable):
        
        if a == 'ConstantMemory' and e == Currentrnbl:
            print(f"[DEBUG] Calling arelements_def.CMCPA_ConstantMemory({currentfolder}, {CurrentSWC_shortname}, {b})")       
            arelements_def.CMCPA_ConstantMemory(currentfolder, CurrentSWC_shortname,b)
            
        elif a == 'PerInstanceParameter'and e == Currentrnbl:
            print(f"[DEBUG] Calling arelements_def.PICPVA_PerInstanceParameter({currentfolder}, {CurrentSWC_shortname}, {b})")           
            arelements_def.PICPVA_PerInstanceParameter(currentfolder, CurrentSWC_shortname,b)

        elif a == 'SharedParameter'and e == Currentrnbl:
            print(f"[DEBUG] Calling arelements_def.SCPVA_SharedParameter({currentfolder}, {CurrentSWC_shortname}, {b})")            
            arelements_def.SCPVA_SharedParameter(currentfolder, CurrentSWC_shortname,b) #ideally we need port parameter here and then shared parameter
            
        elif a == 'ExplicitInterRunnableVariable' and e == Currentrnbl:
            print(f"[DEBUG] Calling arelements_def.IRVRA_ExplicitInterRunnableVariable({b}, {currentfolder}, {CurrentSWC_shortname})")                       
            arelements_def.IRVRA_ExplicitInterRunnableVariable(b, currentfolder, CurrentSWC_shortname )
            
        elif a == 'ImplicitInterRunnableVariables'and e == Currentrnbl:
            print(f"[DEBUG] Calling arelements_def.IRVRA_ImplicitInterRunnableVariable({b}, {currentfolder}, {CurrentSWC_shortname})")           
            arelements_def.IRVRA_ImplicitInterRunnableVariable(b, currentfolder, CurrentSWC_shortname)

        else :
            print(f"[DEBUG] Invalid {a} for IRV access")
    
    # Fetch filtered data from read_write_access() for ReceiverPort and SenderPort
    receiver_port_data = read_write_access("ReceiverPort", Currentrnbl)
    
    sender_port_data = read_write_access("SenderPort", Currentrnbl)
    
    # Debugging: Print filtered data
    print(f"\n[DEBUG] Filtered Data from read_write_access (ReceiverPort, {Currentrnbl}):")
    
    for row in receiver_port_data:
        print(row)
    
    print(f"\n[DEBUG] Filtered Data from read_write_access (SenderPort, {Currentrnbl}):")
    
    for row in sender_port_data:
        print(row)
    
    # Iterate over filtered data for ReceiverPort
    for _, port_name, interface_type, interface_name, data_element, argument in receiver_port_data:
       
        if interface_type == "SenderReceiverInterface" :
           
            if argument == "dra":
                print(f"[DEBUG] Calling arelements_def.DRA_RPort_SR_DataElement({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DRA_RPort_SR_DataElement(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            elif argument == "drpa":
                print(f"[DEBUG] Calling arelements_def.DRPA_RPort_SR_DataElement({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DRPA_RPort_SR_DataElement(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            elif argument == "drpv":
                print(f"[DEBUG] Calling arelements_def.DRPV_RPort_SR_DataElement({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DRPV_RPort_SR_DataElement(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            else :
                print(f"[DEBUG] Invalid {interface_type} for data access dra, drpa and drpv")
        
        elif interface_type == "NvDataInterface":
           
            if argument == "dra":
                print(f"[DEBUG] Calling arelements_def.DRA_RPort_nvd_NvData({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DRA_RPort_nvd_NvData(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            elif argument == "drpa":
                print(f"[DEBUG] Calling arelements_def.DRPA_RPort_nvd_NvData({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DRPA_RPort_nvd_NvData(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            elif argument == "drpv":
                print(f"[DEBUG] Calling arelements_def.DRPV_RPort_nvd_NvData({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DRPV_RPort_nvd_NvData(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            else :
                print(f"[DEBUG] Invalid {interface_type} for data access dra, drpa and drpv")

        elif interface_type == "ParameterInterface":
            print(f"[DEBUG] Calling arelements_def.CPA_RPort_prm_Parameter({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
            arelements_def.CPA_RPort_prm_Parameter(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
        
        elif interface_type == "ClientServerInterface":
            print(f"[DEBUG] Calling arelements_def.SSCP_RPort_CS_Operation({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
            arelements_def.SSCP_RPort_CS_Operation(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
        
        else :
            print(f"[DEBUG] Invalid {interface_type} and port type for data access")

    # Iterate over filtered data for SenderPort
    for _, port_name, interface_type, interface_name, data_element, argument in sender_port_data:
        
        if interface_type == "SenderReceiverInterface" :

            if argument == "dsp":
                print(f"[DEBUG] Calling arelements_def.DSP_RPort_SR_DataElement({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DSP_PPort_SR_DataElement(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            elif argument == "dwa":
                print(f"[DEBUG] Calling arelements_def.DWA_RPort_SR_DataElement({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DWA_PPort_SR_DataElement(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            else :
                print(f"[DEBUG] Invalid {interface_type} for data access dsp,dwa")       
        
        elif interface_type == "NvDataInterface":

            if argument == "dsp":
                print(f"[DEBUG] Calling arelements_def.DSP_RPort_SR_DataElement({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DSP_PPort_SR_DataElement(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            elif argument == "dwa":
                print(f"[DEBUG] Calling arelements_def.DWA_RPort_SR_DataElement({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
                arelements_def.DWA_PPort_SR_DataElement(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
            else :
                print(f"[DEBUG] Invalid {interface_type} for data access dsp,dwa")       
        
        elif interface_type == "ModeSwitchInterface":
            print(f"[DEBUG] Calling arelements_def.MSP_PPort_msi_ModeGroup({currentfolder}, {CurrentSWC_shortname}, {port_name}, {interface_name}, {data_element})")
            arelements_def.MSP_PPort_msi_ModeGroup(currentfolder, CurrentSWC_shortname, port_name, interface_name, data_element)
        
        else :
            print(f"[DEBUG] Invalid {interface_type} and port type for data access")

def rnblaccess_WrittenIRV (Currentrnbl):

    IBVariableType, IBVariableName, ApplicationDataTypeName, Initvalue, AccessingRunnable = excel_reader.read_columns(ib_data, 'B', 'F')

    if any(x in ["ImplicitInterRunnableVariables", "ExplicitInterRunnableVariable"] for x in IBVariableType): #still read or write bifurcation is pending
    
        arelements_def.IRVWA()

        for a,b,c,d,e in zip(IBVariableType, IBVariableName, ApplicationDataTypeName,Initvalue, AccessingRunnable):
            if a == 'ExplicitInterRunnableVariable' and e == Currentrnbl:
                print(f"[DEBUG] Calling arelements_def.IRVWA_ExplicitInterRunnableVariable({b}, {currentfolder}, {CurrentSWC_shortname})")               
                arelements_def.IRVWA_ExplicitInterRunnableVariable(b, currentfolder, CurrentSWC_shortname )
                
            elif a == 'ImplicitInterRunnableVariables'and e == Currentrnbl:
                print(f"[DEBUG] Calling arelements_def.IRVWA_ImplicitInterRunnableVariable({b}, {currentfolder}, {CurrentSWC_shortname})")                
                arelements_def.IRVWA_ImplicitInterRunnableVariable(b, currentfolder, CurrentSWC_shortname)

            else :
                print(f"[DEBUG] Invalid {a} for IRV access")

def read_write_access(port_type_filter, Currentrnbl):
   print("\n[DEBUG] Starting read_write_access function")
   print(f"[DEBUG] Filtering for Port Type: {port_type_filter}, Current Rnbl: {Currentrnbl}")
   # Read required columns from Excel
   port_type_col, port_name_col, interface_type_col, interface_name_col, data_element_col, argument_col, _, accessing_rnbl_col = excel_reader.read_columns(ports, 'B', 'I')
   # Debugging: Print raw data from Excel
   print(f"[DEBUG] Raw Data Read from Excel:")
   print(f"  Port Type Col: {port_type_col}")
   print(f"  Port Name Col: {port_name_col}")
   print(f"  Interface Type Col: {interface_type_col}")
   print(f"  Interface Name Col: {interface_name_col}")
   print(f"  Data Element Col: {data_element_col}")
   print(f"  Argument Col: {argument_col}")
   print(f"  Accessing Rnbl Col: {accessing_rnbl_col}")
   # Ensure all columns are lists and flatten them if necessary
   def flatten(lst):
       return [item for sublist in lst for item in sublist] if any(isinstance(i, list) for i in lst) else lst
   def fill_merged_cells(lst):
       """Fills down values in case of merged cells (None or empty values)."""
       filled_list = []
       last_valid = None
       for item in lst:
           if item:  
               last_valid = item  # Update last valid non-empty value
           filled_list.append(last_valid)  
       return filled_list
   # Apply flattening and merging fixes
   port_type_col = fill_merged_cells(flatten(port_type_col or []))
   port_name_col = fill_merged_cells(flatten(port_name_col or []))
   interface_type_col = fill_merged_cells(flatten(interface_type_col or []))
   interface_name_col = fill_merged_cells(flatten(interface_name_col or []))
   data_element_col = flatten(data_element_col or [])  
   argument_col = flatten(argument_col or [])  
   accessing_rnbl_col = fill_merged_cells(flatten(accessing_rnbl_col or []))
   # Debugging: Print cleaned-up data
   print(f"\n[DEBUG] After Flattening & Merged Cell Handling:")
   print(f"  Port Type Col: {port_type_col}")
   print(f"  Port Name Col: {port_name_col}")
   print(f"  Interface Type Col: {interface_type_col}")
   print(f"  Interface Name Col: {interface_name_col}")
   print(f"  Data Element Col: {data_element_col}")
   print(f"  Argument Col: {argument_col}")
   print(f"  Accessing Rnbl Col: {accessing_rnbl_col}")
   # Ensure mandatory fields are not empty
   if not all(port_name_col) or not all(interface_name_col) or not all(data_element_col) or not all(argument_col):
       print("[ERROR] Mandatory fields cannot be empty!")
       raise ValueError("Mandatory fields (port_name_col, interface_name_col, data_element_col, argument_col) cannot be empty.")
   # Filter rows where accessing_rnbl_col matches Currentrnbl
   filtered_data = []
   for ptype, pname, itype, iname, delem, arg, rnbl in zip(
           port_type_col, port_name_col, interface_type_col, interface_name_col,
           data_element_col, argument_col, accessing_rnbl_col):
       if rnbl == Currentrnbl:
           filtered_data.append((ptype, pname, itype, iname, delem, arg))
   # Debugging: Print filtered data
   print(f"\n[DEBUG] Filtered Data (Matching Currentrnbl={Currentrnbl}):")
   for row in filtered_data:
       print(row)
   # Further filter by port type
   final_filtered_data = [row for row in filtered_data if row[0] == port_type_filter]
   # Debugging: Print final filtered data
   print(f"\n[DEBUG] Final Filtered Data (Matching Port Type: {port_type_filter}):")
   for row in final_filtered_data:
       print(row)
   return final_filtered_data


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  ports __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def Createports():

    
    arelements_def.create_ports(swc_type)


    # Read the port types and names from the specified columns in the ports data
    PortType, PortName, IfType, IfName = excel_reader.read_columns(ports, 'B', 'E')
    
    # Iterate over the port names and their corresponding types
    for port_name, port_type, if_type, if_name in zip(PortName, PortType, IfType, IfName):
        # Check if the port type is 'ReceiverPort'
        if port_type == 'ReceiverPort':
            # Add a new Receiver Port Prototype to the current SWC with the given port name

            if if_type == 'SenderReceiverInterface':
                arelements_def.RPort_SR(port_name,if_name)

            elif if_type == 'ClientServerInterface':
                arelements_def.RPort_CS(port_name,if_name)

            elif if_type == 'ModeSwitchInterface':
                arelements_def.RPort_msi(port_name,if_name)

            elif if_type == 'NvDataInterface':
                arelements_def.RPort_nvd(port_name,if_name)

            elif if_type == 'ParameterInterface':
                arelements_def.RPort_prm(port_name,if_name)
            
            elif if_type == 'TriggerInterface':
                arelements_def.RPort_trigger(port_name,if_name)

            else:
                # Print a message if the port type is unknown
                print(f"Unknown interface type: {if_type} for interface {if_name}")


        # Check if the port type is 'SenderPort'
        elif port_type == 'SenderPort':
            # Add a new Sender Port Prototype to the current SWC with the given port name
            if if_type == 'SenderReceiverInterface':
                arelements_def.PPort_SR(port_name,if_name)

            elif if_type == 'ClientServerInterface':
                arelements_def.PPort_CS(port_name,if_name)

            elif if_type == 'ModeSwitchInterface':
                arelements_def.PPort_msi(port_name,if_name)

            elif if_type == 'NvDataInterface':
                arelements_def.PPort_nvd(port_name,if_name)

            else:
                # Print a message if the port type is unknown
                print(f"Unknown interface type: {if_type} for interface {if_name}")

        else:
            # Print a message if the port type is unknown
            print(f"Unknown port type: {port_type} for port {port_name}")

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  interfaces __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

ApplicationDataTypeName, ApplicationDataTypeCategory, CompuMethodName, CompuMethodCategory, CompuScaleOROffset, EnumStatesORLSB, Unit, DataConstraintName = excel_reader.read_columns(adt_primitive, 'B', 'I')


def createSharedInterfaces():
   """
   Creates shared interfaces based on data from an Excel sheet.
   This function reads from an Excel sheet, groups interfaces by type,
   and creates instances of corresponding interface classes.
   """
   # Read the columns from Excel
   if_type_col, IF_name_col, DE_col, Argument_col, ADt = excel_reader.read_columns(ports, 'D', 'H')

   # Initialize a defaultdict to store interfaces by type
   interface_collections = defaultdict(list)
   # Initialize variables outside the loop
   current_interface_type = None
   current_interface_name = None
   current_data_elements = []
   current_adts = []
   current_arguments = []
   # Iterate through the columns
   for interface_type, interface_name, data_element, argument, adt in zip(if_type_col, IF_name_col, DE_col, Argument_col, ADt):
       # Ensure proper string handling and avoid None values
       interface_type = interface_type.strip() if isinstance(interface_type, str) else None
       interface_name = interface_name.strip() if isinstance(interface_name, str) else None
       data_element = data_element.strip() if isinstance(data_element, str) else None
       argument = argument.strip() if isinstance(argument, str) else None
       adt = adt.strip() if isinstance(adt, str) else None
       # Skip rows with missing essential information
       if not interface_type or not interface_name:
           continue
       # Check if we are still on the same interface_name
       if interface_name == current_interface_name:
           # Append non-empty data elements, ADTs, and arguments
           if data_element:
               current_data_elements.append(data_element)
           if adt:
               current_adts.append(adt)
           if argument:
               current_arguments.append(argument)
       else:
           # If it's a new interface, store the previous one
           if current_interface_name and current_interface_type:
               interface_collections[current_interface_type].append({
                   "name": current_interface_name,
                   "data_elements": current_data_elements,
                   "adts": current_adts,
                   "arguments": current_arguments
               })
           # Start a new interface
           current_interface_type = interface_type
           current_interface_name = interface_name
           current_data_elements = [data_element] if data_element else []
           current_adts = [adt] if adt else []
           current_arguments = [argument] if argument else []
   # Add the last interface to the collection
   if current_interface_name and current_interface_type:
       interface_collections[current_interface_type].append({
           "name": current_interface_name,
           "data_elements": current_data_elements,
           "adts": current_adts,
           "arguments": current_arguments
       })
   # Interface type mapping to constructors
   interface_creation_map = {
       "SenderReceiverInterface": SenderReceiverInterface,
       "NvDataInterface": NvDataInterface,
       "ParameterInterface": ParameterInterface,
       "ModeSwitchInterface": ModeSwitchInterface,
       "ClientServerInterface": ClientServerInterface,
       "TriggerInterface": TriggerInterface
   }
   # Create instances for each interface type
   for interface_type, interfaces in interface_collections.items():
       if interface_type in interface_creation_map:
           creation_func = interface_creation_map[interface_type]
           for interface in interfaces:
               try:
                   # Pass all required arguments
                   creation_func(interface["name"], interface["data_elements"], interface["arguments"], interface["adts"])
               except (TypeError, ValueError, KeyError) as e:
                   print(f"Error creating {interface_type} {interface['name']}: {e}")

def SenderReceiverInterface(currentIF_name, DataElements, Arguments, current_Adt):
    """
    Creates a SenderReceiverInterface object.

    Args:
        currentIF_name: The name of the interface.
        DataElements: A list of data elements associated with the interface.
        current_Adt: A list of ADTs associated with the interface.

    Returns:
        None 
    """
    try:
        # Fetch folder elements for the interface
        SenderReceiver_folder_elements = arxml_structure.get_variable('SenderReceiver_folder_elements')

        # Define the interface using the fetched folder elements
        arelements_def.SenderReceiverInterface(SenderReceiver_folder_elements, currentIF_name)

        # Define the data elements for the interface
        arelements_def.SenderReceiverInterface_DE()

        # Loop through data elements and their corresponding ADTs, and create VDPs
        for itsDE, itsAdt in zip(DataElements, current_Adt):
            arelements_def.SenderReceiverInterface_VDP(itsDE, itsAdt)
    
    except Exception as e:
        print(f"Error creating SenderReceiverInterface for {currentIF_name}: {e}")

def NvDataInterface(currentIF_name, DataElements, Arguments, current_Adt):
    """
    Creates a NvDataInterface object.

    Args:
        currentIF_name: The name of the interface.
        DataElements: A list of data elements associated with the interface.
        current_Adt: A list of ADTs associated with the interface.

    Returns:
        None 
    """
    try:
        # Fetch folder elements for the interface
        NvData_folder_elements = arxml_structure.get_variable('NvData_folder_elements')

        # Define the interface using the fetched folder elements
        arelements_def.NvDataInterface(NvData_folder_elements, currentIF_name)

        # Define the data elements for the interface
        arelements_def.NvDataInterface_DE()

        # Loop through data elements and their corresponding ADTs, and create VDPs
        for itsDE, itsAdt in zip(DataElements, current_Adt):
            arelements_def.NvDataInterface_VDP(itsDE, itsAdt)
    
    except Exception as e:
        print(f"Error creating NvDataInterface for {currentIF_name}: {e}")

def ParameterInterface(currentIF_name, DataElements, Arguments, current_Adt):
    """
    Creates a ParameterInterface object.

    Args:
        currentIF_name: The name of the interface.
        DataElements: A list of data elements associated with the interface.
        current_Adt: A list of ADTs associated with the interface.

    Returns:
        None 
    """
    try:
        # Fetch folder elements for the interface
        Parameter_folder_elements = arxml_structure.get_variable('Parameter_folder_elements')

        # Define the interface using the fetched folder elements
        arelements_def.ParameterInterface(Parameter_folder_elements, currentIF_name)

        # Define the data elements for the interface
        arelements_def.ParameterInterface_DE()

        # Loop through data elements and their corresponding ADTs, and create VDPs
        for itsDE, itsAdt in zip(DataElements, current_Adt):
            arelements_def.ParameterInterface_VDP(itsDE, itsAdt)
    
    except Exception as e:
        print(f"Error creating ParameterInterface for {currentIF_name}: {e}")

def ModeSwitchInterface(currentIF_name, DataElements, Arguments, current_Adt):
    """
    Creates a ModeSwitchInterface object.

    Args:
        currentIF_name: The name of the interface.
        DataElements: The name of operation name
        Arguments: The name of arguments
        current_Adt: A list of ADTs associated with the arguments.
   
    Returns:
        None 
    """
    ModeDeclarationGroup_shortname = DataElements[0]

    mode_Category = 'ALPHABETIC_ORDER'  #default and other is EXPLICIT_ORDER

    Init_Mode = Arguments[0]

    try:
        # Fetch folder elements for the interface
        ModeSwitch_folder_elements = arxml_structure.get_variable('ModeSwitch_folder_elements')

        # Define the interface using the fetched folder elements
        arelements_def.ModeDeclarationGroup(ModeSwitch_folder_elements, ModeDeclarationGroup_shortname, mode_Category,Init_Mode)

        # Loop through data elements and their corresponding ADTs, and create VDPs
        for ModeDeclaration_shortname in Arguments :
            arelements_def.ModeDeclarationGroup_Exp(ModeDeclaration_shortname)

        # Define the data elements for the interface
        arelements_def.ModeSwitchInterface(ModeSwitch_folder_elements, currentIF_name, ModeDeclarationGroup_shortname)
    
    except Exception as e:
        print(f"Error creating ModeSwitchInterface for {currentIF_name}: {e}")

def ClientServerInterface(currentIF_name, DataElements, Arguments, current_Adt):
    """
    Creates a ClientServerInterface object.

    Args:
        currentIF_name: The name of the interface.
        DataElements: The name of ModeDeclarationGroup
        Arguments: The name of modes from ModeDeclarationGroup
   
    Returns:
        None 
    """   

    try:
        # Fetch folder elements for the interface
        ClientServer_folder_elements = arxml_structure.get_variable('ClientServer_folder_elements')
        # Define the interface
        arelements_def.ClientServerInterface(ClientServer_folder_elements, currentIF_name)
        arelements_def.ClientServerInterface_Opr()
        # Sort the data based on Operation_shortname to ensure correct grouping
        sorted_data = sorted(zip(DataElements, Arguments, current_Adt), key=lambda x: x[0])
        # Group by `Operation_shortname`
        for Operation_shortname, group in groupby(sorted_data, key=lambda x: x[0]):
            arelements_def.ClientServerInterface_CSOpr(Operation_shortname)  # Called once per operation
            arelements_def.ClientServerInterface_Args()  # Called once per operation
            # Iterate over all `itsArg` and `itsAdt` associated with the same `Operation_shortname`
            for _, itsArg, itsAdt in group:
                arelements_def.ClientServerInterface_Arg(itsArg, itsAdt)  # Called for each argument
    except Exception as e:
        print(f"Error creating ClientServerInterface for {currentIF_name}: {e}")

def TriggerInterface(currentIF_name, DataElements, Arguments, current_Adt):
    """
    Creates a TriggerInterface object.

    Args:
        currentIF_name: The name of the interface.
        DataElements: A list of triggers associated with the interface.
        Arguments: A list of cse code and cse code factor associated with the interface.

    Returns:
        None 
    """
    try:
        # Fetch folder elements for the interface
        Trigger_folder_elements = arxml_structure.get_variable('Trigger_folder_elements')

        # Define the interface using the fetched folder elements
        arelements_def.TriggerInterface(Trigger_folder_elements, currentIF_name)

        arelements_def.TriggerInterface_trigs()

        arg_iter = iter(Arguments)  # Create an iterator for Arguments
        for trigger_shortname in DataElements[::2]:  # Pick every alternate trigger_shortname
            cse_code = next(arg_iter, None)  # Get the first argument value
            cse_code_factor = next(arg_iter, None)  # Get the second argument value
            if cse_code is not None and cse_code_factor is not None:
                arelements_def.TriggerInterface_trig(trigger_shortname, cse_code, cse_code_factor)
    
    except Exception as e:
        print(f"Error creating TriggerInterface for {currentIF_name}: {e}")


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  compu method __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def handle_identical(currentcompumethod, NA, NA1, Current_Unit):
   if NA is None or NA == "":
       print("CompuMethodInfo is not applicable for IDENTICAL category.")
   if NA1 is None or NA1 == "":
       print("CompuMethodInfo is not applicable for IDENTICAL category.")
   if Current_Unit is None or Current_Unit == "":
       print("Unit is not applicable for IDENTICAL category.")
   arelements_def.CompuMethod_IDENTICAL(CompuMethods_shared_folder_elements, currentcompumethod, Current_Unit)

def handle_texttable(currentcompumethod, Compu_Scale, Enum_States, Current_Unit):

    # Check if Compu_Scale is None or an empty string
    if Compu_Scale is None or Compu_Scale == "":
        print("CompuMethodInfo is not applicable for TEXTTABLE category.")
    # Check if Enum_States is None or an empty string
    if Enum_States is None or Enum_States == "":
        print("CompuMethodInfo is not applicable for TEXTTABLE category.")
    # Check if Current_Unit is None or an empty string
    if Current_Unit is None or Current_Unit == "":
        print("Unit is not applicable for TEXTTABLE category.")

 
    arelements_def.CompuMethod_text(CompuMethods_shared_folder_elements, currentcompumethod, Current_Unit)


    # Iterate over Compu_Scale and Enum_States to call the function for each pair
    for scale, state in zip(Compu_Scale, Enum_States):
        arelements_def.text_compu_Scale(scale, state)

    # Set default value if it's not already set
    defaultValue = '0' 
    if defaultValue is not None :
        arelements_def.text_compu_DefaultValue(defaultValue)

def handle_linear(currentcompumethod, CompuScaleOROffset, EnumStatesORLSB, Current_Unit):

    if CompuScaleOROffset is None or CompuScaleOROffset == "":
        print("CompuMethodInfo is not applicable for LINEAR category.")
    if EnumStatesORLSB is None or EnumStatesORLSB == "":
        print("CompuMethodInfo is not applicable for LINEAR category.")        
    if Current_Unit is None or Current_Unit == "":
        print("Unit is not applicable for LINEAR category.")

    arelements_def.CompuMethod_linear(CompuMethods_shared_folder_elements, currentcompumethod, Current_Unit)

    num_a = CompuScaleOROffset[0]
    num_b = CompuScaleOROffset[1]
    den_a = EnumStatesORLSB[0]

    arelements_def.linear_compu_scale(num_a,num_b,den_a)

def handle_scale_linear(compu_method_info, unit):
   # function def not available in arelement def
   if compu_method_info is None or compu_method_info == "":
       print("CompuMethodInfo is not applicable for SCALE_LINEAR category.")
   if unit is None or unit == "":
       print("Unit is not applicable for SCALE_LINEAR category.")

def handle_scale_linear_and_texttable(compu_method_info, unit):
   if compu_method_info is None or compu_method_info == "":
       print("CompuMethodInfo is not applicable for SCALE_LINEAR_AND_TEXTTABLE category.")
   if unit is None or unit == "":
       print("Unit is not applicable for SCALE_LINEAR_AND_TEXTTABLE category.")

def handle_rat_func(currentcompumethod, CompuScaleOROffset, EnumStatesORLSB, Current_Unit):

    if CompuScaleOROffset is None or CompuScaleOROffset == "":
        print("CompuMethodInfo is not applicable for LINEAR category.")
    if EnumStatesORLSB is None or EnumStatesORLSB == "":
        print("CompuMethodInfo is not applicable for LINEAR category.")        
    if Current_Unit is None or Current_Unit == "":
        print("Unit is not applicable for LINEAR category.")

    arelements_def.CompuMethod_rat_func(CompuMethods_shared_folder_elements, currentcompumethod, Current_Unit)

    num_a = CompuScaleOROffset[0]
    num_b = CompuScaleOROffset[1]
    num_c = CompuScaleOROffset[2]
    den_a = EnumStatesORLSB[0]
    den_b = EnumStatesORLSB[1]
    den_c = EnumStatesORLSB[2]

    arelements_def.rat_func_compu_scale(num_a,num_b,num_c,den_a,den_b,den_c)

def handle_scale_rat_func(currentcompumethod, CompuScaleOROffset, EnumStatesORLSB, Current_Unit):

    if CompuScaleOROffset is None or CompuScaleOROffset == "":
        print("CompuMethodInfo is not applicable for LINEAR category.")
    if EnumStatesORLSB is None or EnumStatesORLSB == "":
        print("CompuMethodInfo is not applicable for LINEAR category.")        
    if Current_Unit is None or Current_Unit == "":
        print("Unit is not applicable for LINEAR category.")

    arelements_def.CompuMethod_Scale_rat_text(CompuMethods_shared_folder_elements, currentcompumethod, Current_Unit)

    ll = CompuScaleOROffset[0]
    num_a = CompuScaleOROffset[1]
    num_b = CompuScaleOROffset[2]
    num_c = CompuScaleOROffset[3]
    ul = EnumStatesORLSB[0]
    den_a = EnumStatesORLSB[1]
    den_b = EnumStatesORLSB[2]
    den_c = EnumStatesORLSB[3]

    arelements_def.Scale_rat_text_compu_scale(ll,ul,num_a,num_b,num_c,den_a,den_b,den_c)

    # Set default value if it's not already set
    defaultValue = '0' 
    if defaultValue is not None :
        arelements_def.Scale_rat_text_compu_default_value(defaultValue)

def handle_scale_rational_and_texttable(compu_method_info, unit):
   if compu_method_info is None or compu_method_info == "":
       print("CompuMethodInfo is not applicable for SCALE_RATIONAL_AND_TEXTTABLE category.")
   if unit is None or unit == "":
       print("Unit is not applicable for SCALE_RATIONAL_AND_TEXTTABLE category.")

def handle_tab_nointp(currentcompumethod, Compu_Scale, Enum_States, Current_Unit):


    # Check if Compu_Scale is None or an empty string
    if Compu_Scale is None or Compu_Scale == "":
        print("CompuMethodInfo is not applicable for TEXTTABLE category.")
    # Check if Enum_States is None or an empty string
    if Enum_States is None or Enum_States == "":
        print("CompuMethodInfo is not applicable for TEXTTABLE category.")
    # Check if Current_Unit is None or an empty string
    if Current_Unit is None or Current_Unit == "":
        print("Unit is not applicable for TEXTTABLE category.")

 
    arelements_def.CompuMethod_tab_nointp(CompuMethods_shared_folder_elements, currentcompumethod, Current_Unit)

    # CompuMethod_tab_nointp(CompuMethods_shared_folder_elements, compu_method_shortname, unit)


    # Iterate over Compu_Scale and Enum_States to call the function for each pair
    for scale, state in zip(Compu_Scale, Enum_States):
        arelements_def.tab_nointp_compu_Scale(scale, state)
        # tab_nointp_compu_Scale(value, enum)

    # Set default value if it's not already set
    defaultValue = '0' 
    if defaultValue is not None :
        arelements_def.tab_nointp_compu_Scale_DefaultValue(defaultValue)

def handle_bitfield_texttable(compu_method_info, unit):
   if compu_method_info is None or compu_method_info == "":
       print("CompuMethodInfo is not applicable for BITFIELD_TEXTTABLE category.")
   if unit is None or unit == "":
       print("Unit is not applicable for BITFIELD_TEXTTABLE category.")

def createcompumethod():
   global CompuMethods_shared_folder_elements
   CompuMethods_shared_folder_elements = arxml_structure.get_variable('CompuMethods_shared_folder_elements')
   # Read columns separately before zipping
   CompuMethodName, CompuMethodCategory, CompuScaleOROffset, EnumStatesORLSB, Unit = excel_reader.read_columns(
       adt_primitive, 'D', 'H'
   )
   # Debug: Print raw data from Excel
   print("\nDEBUG: Raw data from Excel")
   print(f"CompuMethodName: {CompuMethodName}")
   print(f"CompuMethodCategory: {CompuMethodCategory}")
   print(f"CompuScaleOROffset: {CompuScaleOROffset}")
   print(f"EnumStatesORLSB: {EnumStatesORLSB}")
   print(f"Unit: {Unit}")
   # Dictionary to store collected data
   compu_methods = {}
   current_compu_method_name = None
   current_compu_method_category = None
   current_unit = None
   compu_scale = []
   enum_states = []
   for name, category, scale, enum, unit in zip(CompuMethodName, CompuMethodCategory, CompuScaleOROffset, EnumStatesORLSB, Unit):
       if name and name != current_compu_method_name:  
           # **Process the previous collected method before switching**
           if current_compu_method_name is not None:
               compu_methods[current_compu_method_name] = {
                   "category": current_compu_method_category if current_compu_method_category else "IDENTICAL",
                   "unit": current_unit if current_unit else None,
                   "compu_scale": compu_scale if compu_scale else [],
                   "enum_states": enum_states if enum_states else [],
               }
               print(f"DEBUG: Saved data for {current_compu_method_name} -> Scale: {compu_scale}, Enum: {enum_states}")
           # **Reset lists and assign new method details**
           current_compu_method_name = name
           current_compu_method_category = category
           current_unit = unit
           compu_scale = []  # Reset only after processing previous values
           enum_states = []
       # **Append values for the current method**
       if scale is not None:
           compu_scale.append(scale)
       if enum is not None:
           enum_states.append(enum)
   # **Save the last collected method**
   if current_compu_method_name:
       compu_methods[current_compu_method_name] = {
           "category": current_compu_method_category if current_compu_method_category else "IDENTICAL",
           "unit": current_unit if current_unit else None,
           "compu_scale": compu_scale if compu_scale else [],
           "enum_states": enum_states if enum_states else [],
       }
       print(f"DEBUG: Final Save for {current_compu_method_name} -> Scale: {compu_scale}, Enum: {enum_states}")
   print("\nDEBUG: Completed CompuMethod collection, now processing...")
   # **Process all collected methods using if-else**
   for method_name, details in compu_methods.items():
       category = details["category"]
       compu_scale = details["compu_scale"]
       enum_states = details["enum_states"]
       unit = details["unit"]
       print(f"\nDEBUG: Processing CompuMethod: {method_name}, Category: {category}, Scale: {compu_scale}, Enum: {enum_states}, Unit: {unit}")
       if category == "IDENTICAL":
           handle_identical(method_name, compu_scale, enum_states, unit)
       elif category == "TEXTTABLE":
           handle_texttable(method_name, compu_scale, enum_states, unit)
       elif category == "LINEAR":
           handle_linear(method_name, compu_scale, enum_states, unit)
       elif category == "SCALE_LINEAR":
           handle_scale_linear(method_name, compu_scale, enum_states, unit)
       elif category == "SCALE_LINEAR_AND_TEXTTABLE":
           handle_scale_linear_and_texttable(method_name, compu_scale, enum_states, unit)
       elif category == "RAT_FUNC":
           handle_rat_func(method_name, compu_scale, enum_states, unit)
       elif category == "SCALE_RAT_FUNC":
           handle_scale_rat_func(method_name, compu_scale, enum_states, unit)
       elif category == "SCALE_RATIONAL_AND_TEXTTABLE":
           handle_scale_rational_and_texttable(method_name, compu_scale, enum_states, unit)
       elif category == "TAB_NOINTP":
           handle_tab_nointp(method_name, compu_scale, enum_states, unit)
       elif category == "BITFIELD_TEXTTABLE":
           handle_bitfield_texttable(method_name, compu_scale, enum_states, unit)
       else:
           print(f"Warning: Invalid CompuMethod category '{category}' for method '{method_name}'")
   print("\nDEBUG: CompuMethod processing completed.")


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  data constraint __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def createDC():
    # Read columns from the adt_primitive data source
    DataConstraintName, DataConstraintType, Min, Max = excel_reader.read_columns(adt_primitive, 'I', 'L')
    
    # Get the Data Constraints package from the shared elements
    DataConstr_folder_elements = arxml_structure.get_variable('DataConstr_folder_elements')
    
    # Iterate through the data constraints
    for a, b, c, d in zip(DataConstraintName, DataConstraintType, Min, Max):
        constrLevel = 0  # Initialize constraint level
        lowerLimit = c   # Set lower limit from the read data
        upperLimit = d   # Set upper limit from the read data
        
        # Check if the constraint type is 'internalConstrs'
        if b == 'internalConstrs':
            arelements_def.DataConstr_Int(DataConstr_folder_elements,a,lowerLimit,upperLimit)
        
        # Check if the constraint type is 'physConstrs'
        elif b == 'physConstrs':
            # Create a new physical data constraint
            arelements_def.DataConstr_phy(DataConstr_folder_elements,a,lowerLimit,upperLimit)
        
        # If the constraint type is neither 'internalConstrs' nor 'physConstrs'
        else:
            if b is not None:
                print(f"Value of b: {b}")  # Print the value of b for debugging purposes

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  ADT __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def createprimitive():
   # Read columns from the adt_primitive data source
   APDT_name, APDT_category, APDT_CMname, _, _, _, APDT_unit, APDT_DCname = excel_reader.read_columns(adt_primitive, 'B', 'I')
   # Get the Data Constraints package from shared elements
   Primitive_folder_elements = arxml_structure.get_variable('Primitive_folder_elements')
   # Define category-function mapping for scalability
   category_mapping = {
       'Value': arelements_def.ApplicationPrimitiveDataType_Val,
       'BOOLEAN': arelements_def.Bool_ApplicationPrimitiveDataType,
       'STRING': arelements_def.String_ApplicationPrimitiveDataType
   }
   processed_names = set()  # Track processed names to avoid duplicates
   # Iterate through the data constraints
   for name, category, cm_name, unit, dc_name in zip(APDT_name, APDT_category, APDT_CMname, APDT_unit, APDT_DCname):
       # Skip empty or duplicate rows
       if not name or not category:
           config.logging.warning(f"Skipping row due to missing required fields: {name}, {category}")
           continue
       if name in processed_names:
           config.logging.info(f"Skipping duplicate entry: {name}")
           continue
       # Mark as processed
       processed_names.add(name)
       # Process based on category
       func = category_mapping.get(category)
       if func:
           func(Primitive_folder_elements, name, cm_name, dc_name, unit)
       else:
           config.logging.warning(f"Category '{category}' is not supported")

def createcomposite():
   """Creates composite data types (Record & Array) from Excel data."""
   print("Reading Excel columns...")
   Composite_category, ARDT_ShortName, ARDT_element_shortname, ARDT_element_type, data_type = excel_reader.read_columns( adt_composite, 'B', 'F' )
   print(f"Read {len(Composite_category)} rows from Excel.")
   previous_category = None
   previous_shortname = None
   record_elements = []
   for idx, (a, b, c, d, e) in enumerate(zip( Composite_category, ARDT_ShortName, ARDT_element_shortname, ARDT_element_type, data_type )):
       print(f"\nProcessing Row {idx + 1}:")
       print(f"Category: {a}, ShortName: {b}, ElementShortName: {c}, Type: {d}, DataType: {e}")
       # Handling RECORD (STRUCTURE)
       if a == 'RECORD':  
           print("Detected RECORD category.")
           if previous_category == 'RECORD' and previous_shortname and previous_shortname != b:
               print(f"Creating ApplicationRecordDataType for {previous_shortname}")
               Record_folder_elements = arxml_structure.get_variable('Record_folder_elements')
               arelements_def.ApplicationRecordDataType(Record_folder_elements, previous_shortname)
               print(f"Record elements before unpacking: {record_elements}")  # Debugging print
               for elem in record_elements:
                   if len(elem) == 3:
                       elem_c, elem_d, elem_e = elem
                       print(f"Adding Record Element: {elem_c}, Type: {elem_d}, DataType: {elem_e}")
                       arelements_def.ApplicationRecordDataType_elements(elem_c, elem_d, elem_e)
                   else:
                       print(f"Skipping malformed record element: {elem}")  # If something is incorrect
               record_elements = []  
               print("Cleared record elements for new RECORD.")
           previous_category = a
           previous_shortname = b
           record_elements.append((c, d, e))
           print(f"Collected RECORD element: {c}, Type: {d}, DataType: {e}")
       # Handling ARRAY
       elif a == 'ARRAY':  
           print("Detected ARRAY category.")
           Array_folder_elements = arxml_structure.get_variable('Array_folder_elements')
           if (b, c, d, e) not in record_elements:
               if d == 'VARIABLE':
                   print(f"Creating ApplicationArrayDataType_Variable for {b}")
                   arelements_def.ApplicationArrayDataType_Variable(Array_folder_elements, b, e, c)
               elif d == 'FIXED':
                   print(f"Creating ApplicationArrayDataType_Fixed for {b}")
                   arelements_def.ApplicationArrayDataType_Fixed(Array_folder_elements, b, e, c)
               else:
                   print(f"Invalid array category '{d}' for '{b}'. Expected 'VARIABLE' or 'FIXED'.")
               record_elements.append((b, c, d, e))
               print(f"Marked ARRAY row as processed: {b}, {c}, {d}, {e}")
           else:
               print("Skipping duplicate ARRAY row.")
   # Finalizing the last RECORD
   if previous_category == 'RECORD' and previous_shortname:
       print(f"Finalizing last RECORD: {previous_shortname}")
       Record_folder_elements = arxml_structure.get_variable('Record_folder_elements')
       arelements_def.ApplicationRecordDataType(Record_folder_elements, previous_shortname)
       print(f"Record elements before unpacking: {record_elements}")  # Debugging print
       for elem in record_elements:
           if len(elem) == 3:
               elem_c, elem_d, elem_e = elem
               print(f"Adding final Record Element: {elem_c}, Type: {elem_d}, DataType: {elem_e}")
               arelements_def.ApplicationRecordDataType_elements(elem_c, elem_d, elem_e)
           else:
               print(f"Skipping malformed record element: {elem}")
   print("Processing completed.")


#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  IDT __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def createcustomIDT():
   # Get the Data Constraints package from shared elements
   ImplementationDataTypes_folder_elements = arxml_structure.get_variable('ImplementationDataTypes_folder_elements')
   print("[DEBUG] Retrieved ImplementationDataTypes_folder_elements")
   # Read columns: type, shortname, arraysize/idtelementshortname, IDT
   IDT_type, IDT_shortname, IRDT_element_shortname, data_type = excel_reader.read_columns(idt, 'B', 'E')
   print(f"[DEBUG] Read {len(IDT_type)} IDT entries from Excel")
   record_shortname = None
   record_elements = []
   for i in range(len(IDT_type)):
       idt_type = IDT_type[i]
       idt_shortname = IDT_shortname[i]
       idt_element_shortname = IRDT_element_shortname[i]
       idt_data_type = data_type[i]
       print(f"[DEBUG] Processing index {i}: Type={idt_type}, Shortname={idt_shortname}, Element={idt_element_shortname}, DataType={idt_data_type}")
       if idt_type == 'ARRAY_FIXED':
           # Process any pending RECORD before handling ARRAY_FIXED
           if record_shortname:
               print(f"[DEBUG] Processing collected RECORD: {record_shortname} with {len(record_elements)} elements")
               arelements_def.ImplementationDataType_Structure(ImplementationDataTypes_folder_elements, record_shortname)
               for element_shortname, element_data_type in record_elements:
                   print(f"[DEBUG] Adding RECORD element: {element_shortname}, {element_data_type}")
                   arelements_def.ImplementationDataType_Record_elements(element_shortname, element_data_type)
               record_shortname = None
               record_elements.clear()
           print(f"[DEBUG] Creating ARRAY_FIXED: {idt_shortname}, {idt_element_shortname}, {idt_data_type}")
           arelements_def.ImplementationDataType_ArrayFixed(ImplementationDataTypes_folder_elements,
                                                             idt_shortname, idt_element_shortname, idt_data_type)
       elif idt_type == 'ARRAY_VARIABLE':
           # Process any pending RECORD before handling ARRAY_VARIABLE
           if record_shortname:
               print(f"[DEBUG] Processing collected RECORD: {record_shortname} with {len(record_elements)} elements")
               arelements_def.ImplementationDataType_Structure(ImplementationDataTypes_folder_elements, record_shortname)
               for element_shortname, element_data_type in record_elements:
                   print(f"[DEBUG] Adding RECORD element: {element_shortname}, {element_data_type}")
                   arelements_def.ImplementationDataType_Record_elements(element_shortname, element_data_type)
               record_shortname = None
               record_elements.clear()
           print(f"[DEBUG] Creating ARRAY_VARIABLE: {idt_shortname}, {idt_element_shortname}, {idt_data_type}")
           arelements_def.ImplementationDataType_ArrayVariable(ImplementationDataTypes_folder_elements,
                                                                idt_shortname, idt_element_shortname, idt_data_type)
       elif idt_type == 'PRIMITIVE':
           # Process any pending RECORD before handling PRIMITIVE
           if record_shortname:
               print(f"[DEBUG] Processing collected RECORD: {record_shortname} with {len(record_elements)} elements")
               arelements_def.ImplementationDataType_Structure(ImplementationDataTypes_folder_elements, record_shortname)
               for element_shortname, element_data_type in record_elements:
                   print(f"[DEBUG] Adding RECORD element: {element_shortname}, {element_data_type}")
                   arelements_def.ImplementationDataType_Record_elements(element_shortname, element_data_type)
               record_shortname = None
               record_elements.clear()
           print(f"[DEBUG] Creating PRIMITIVE: {idt_shortname}, {idt_data_type}")
           arelements_def.ImplementationDataType(ImplementationDataTypes_folder_elements,
                                                 idt_shortname, idt_data_type)
       elif idt_type == 'RECORD':
           if record_shortname and idt_shortname != record_shortname:
               # If a new RECORD structure starts, process the previous one
               print(f"[DEBUG] Processing collected RECORD: {record_shortname} with {len(record_elements)} elements")
               arelements_def.ImplementationDataType_Structure(ImplementationDataTypes_folder_elements, record_shortname)
               for element_shortname, element_data_type in record_elements:
                   print(f"[DEBUG] Adding RECORD element: {element_shortname}, {element_data_type}")
                   arelements_def.ImplementationDataType_Record_elements(element_shortname, element_data_type)
               record_elements.clear()
           record_shortname = idt_shortname
           record_elements.append((idt_element_shortname, idt_data_type))
           print(f"[DEBUG] Collecting RECORD elements for {record_shortname}: {idt_element_shortname}, {idt_data_type}")
       else:
           print(f"[ERROR] Unknown IDT_type: {idt_type}")
           raise ValueError(f"Unknown IDT_type: {idt_type}")
   # Process any remaining RECORD at the end of the loop
   if record_shortname and record_elements:
       print(f"[DEBUG] Final Processing for RECORD: {record_shortname} with {len(record_elements)} elements")
       arelements_def.ImplementationDataType_Structure(ImplementationDataTypes_folder_elements, record_shortname)
       for element_shortname, element_data_type in record_elements:
           print(f"[DEBUG] Adding RECORD element: {element_shortname}, {element_data_type}")
           arelements_def.ImplementationDataType_Record_elements(element_shortname, element_data_type)
   print("[DEBUG] Finished processing all IDT entries")







#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#               # ####################### =============== %%%%%%%%% ++++++++++++ ---------- __________ SECTION :  main __________ ----------  ++++++++++++ %%%%%%%%% =============== ####################### #               #
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

def Main():
    # Execute the main sequence of functions for project setup

    CreateSwcs()           # Create software components
    createcompumethod()    # Create computation methods
    createDC()             # Create data constraints
    createprimitive()
    createcomposite()
    createcustomIDT()
    createSharedInterfaces()

    tree = ET.ElementTree(root)

    try:
        root1 = tree.getroot()

        # Remove namespaces
        root2 = remove_namespaces(root1)

        # Add indentation
        indent(root2)

        # Write the XML declaration with double quotes
        with open("output113.arxml", "wb") as f:
            f.write(b'<?xml version="1.0" encoding="utf-8"?>\n')
            tree.write(f, encoding="utf-8", xml_declaration=False)

        print(f"Successfully created with proper indentation and XML declaration.")

    except FileNotFoundError as e:
        print(f"Error: {e}. Please enter a valid ARXML file path.")

# Entry point of the script
if __name__ == "__main__":
    Main()  # Run the main function
    print("Ready")  # Indicate that the process is complete
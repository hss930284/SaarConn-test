import xml.etree.ElementTree as ET
import xml.dom.minidom
import rng
from data_type_utils import DataProcessor  # Import the DataProcessor class 

# Create an instance of the DataProcessor class
processor = DataProcessor()
# Create the root element




root = ET.Element("AUTOSAR", 
                     attrib={
                         "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance",
                         "xmlns": "http://autosar.org/schema/r4.0",
                         "xsi:schemaLocation": "http://autosar.org/schema/r4.0 AUTOSAR_4-0-2.xsd"
                     })


########## application data type ########## 

# 1. APDT 

def ApplicationPrimitiveDataType_Val_Invalid(Primitive_folder_elements, ApplicationPrimitiveDataType_shortname, APDT_CompuMethod, APDT_DataConstr, APDT_unit, APDT_InvalidVal ):#completed
	a=processor.value_to_str(APDT_InvalidVal)
	application_primitive_data_type=ET.SubElement(Primitive_folder_elements,'APPLICATION-PRIMITIVE-DATA-TYPE')
	application_primitive_data_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(application_primitive_data_type,'SHORT-NAME')
	short_name.text=ApplicationPrimitiveDataType_shortname
	category=ET.SubElement(application_primitive_data_type,'CATEGORY')
	category.text='VALUE'
	sw_data_def_props=ET.SubElement(application_primitive_data_type,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text='NOT-ACCESSIBLE'
	compu_method_ref=ET.SubElement(sw_data_def_props_conditional,'COMPU-METHOD-REF')
	compu_method_ref.text=f'/SharedElements/CompuMethods/{APDT_CompuMethod}'
	compu_method_ref.attrib={'DEST':'COMPU-METHOD'}
	data_constr_ref=ET.SubElement(sw_data_def_props_conditional,'DATA-CONSTR-REF')
	data_constr_ref.text=f'/SharedElements/DataConstr/{APDT_DataConstr}'
	data_constr_ref.attrib={'DEST':'DATA-CONSTR'}
	invalid_value=ET.SubElement(sw_data_def_props_conditional,'INVALID-VALUE')
	application_value_specification=ET.SubElement(invalid_value,'APPLICATION-VALUE-SPECIFICATION')
	category=ET.SubElement(application_value_specification,'CATEGORY')
	category.text='VALUE'
	sw_value_cont=ET.SubElement(application_value_specification,'SW-VALUE-CONT')
	sw_values_phys=ET.SubElement(sw_value_cont,'SW-VALUES-PHYS')
	v=ET.SubElement(sw_values_phys,'V')
	v.text=a
	unit_ref=ET.SubElement(sw_data_def_props_conditional,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{APDT_unit}'
	unit_ref.attrib={'DEST':'UNIT'}

def Bool_ApplicationPrimitiveDataType(Primitive_folder_elements, ApplicationPrimitiveDataType_shortname, APDT_CompuMethod, APDT_DataConstr, APDT_unit):#completed

	application_primitive_data_type=ET.SubElement(Primitive_folder_elements,'APPLICATION-PRIMITIVE-DATA-TYPE')
	application_primitive_data_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(application_primitive_data_type,'SHORT-NAME')
	short_name.text=ApplicationPrimitiveDataType_shortname
	category=ET.SubElement(application_primitive_data_type,'CATEGORY')
	category.text='BOOLEAN'
	sw_data_def_props=ET.SubElement(application_primitive_data_type,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text='NOT-ACCESSIBLE'
	compu_method_ref=ET.SubElement(sw_data_def_props_conditional,'COMPU-METHOD-REF')
	compu_method_ref.text=f'/SharedElements/CompuMethods/{APDT_CompuMethod}'
	compu_method_ref.attrib={'DEST':'COMPU-METHOD'}
	data_constr_ref=ET.SubElement(sw_data_def_props_conditional,'DATA-CONSTR-REF')
	data_constr_ref.text=f'/SharedElements/DataConstr/{APDT_DataConstr}'
	data_constr_ref.attrib={'DEST':'DATA-CONSTR'}
	unit_ref=ET.SubElement(sw_data_def_props_conditional,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{APDT_unit}'
	unit_ref.attrib={'DEST':'UNIT'}

def ApplicationPrimitiveDataType_Val(Primitive_folder_elements, ApplicationPrimitiveDataType_shortname, APDT_CompuMethod, APDT_DataConstr, APDT_unit):#completed


	application_primitive_data_type=ET.SubElement(Primitive_folder_elements,'APPLICATION-PRIMITIVE-DATA-TYPE')
	application_primitive_data_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(application_primitive_data_type,'SHORT-NAME')
	short_name.text=ApplicationPrimitiveDataType_shortname
	category=ET.SubElement(application_primitive_data_type,'CATEGORY')
	category.text='VALUE'
	sw_data_def_props=ET.SubElement(application_primitive_data_type,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text='NOT-ACCESSIBLE'
	compu_method_ref=ET.SubElement(sw_data_def_props_conditional,'COMPU-METHOD-REF')
	compu_method_ref.text=f'/SharedElements/CompuMethods/{APDT_CompuMethod}'
	compu_method_ref.attrib={'DEST':'COMPU-METHOD'}
	data_constr_ref=ET.SubElement(sw_data_def_props_conditional,'DATA-CONSTR-REF')
	data_constr_ref.text=f'/SharedElements/DataConstr/{APDT_DataConstr}'
	data_constr_ref.attrib={'DEST':'DATA-CONSTR'}
	unit_ref=ET.SubElement(sw_data_def_props_conditional,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{APDT_unit}'
	unit_ref.attrib={'DEST':'UNIT'}

def String_ApplicationPrimitiveDataType(Primitive_folder_elements,ApplicationPrimitiveDataType_shortname,APDT_unit): #completed

	application_primitive_data_type=ET.SubElement(Primitive_folder_elements,'APPLICATION-PRIMITIVE-DATA-TYPE')
	application_primitive_data_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(application_primitive_data_type,'SHORT-NAME')
	short_name.text=ApplicationPrimitiveDataType_shortname
	category=ET.SubElement(application_primitive_data_type,'CATEGORY')
	category.text='STRING'
	sw_data_def_props=ET.SubElement(application_primitive_data_type,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text='NOT-ACCESSIBLE'
	sw_text_props=ET.SubElement(sw_data_def_props_conditional,'SW-TEXT-PROPS')
	array_size_semantics=ET.SubElement(sw_text_props,'ARRAY-SIZE-SEMANTICS')
	array_size_semantics.text='VARIABLE-SIZE'
	sw_max_text_size=ET.SubElement(sw_text_props,'SW-MAX-TEXT-SIZE')
	sw_max_text_size.text='16'
	unit_ref=ET.SubElement(sw_data_def_props_conditional,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{APDT_unit}'
	unit_ref.attrib={'DEST':'UNIT'}

# 2. ARDT

def ApplicationRecordDataType(Record_folder_elements, ARDT_ShortName):#completed
	global ARDT_elements

	application_record_data_type=ET.SubElement(Record_folder_elements,'APPLICATION-RECORD-DATA-TYPE')
	application_record_data_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(application_record_data_type,'SHORT-NAME')
	short_name.text=ARDT_ShortName
	category=ET.SubElement(application_record_data_type,'CATEGORY')
	category.text='STRUCTURE'
	ARDT_elements=ET.SubElement(application_record_data_type,'ELEMENTS')

def ApplicationRecordDataType_elements(ARDT_element_shortname, ARDT_element_type, data_type):#completed 
	application_record_element=ET.SubElement(ARDT_elements,'APPLICATION-RECORD-ELEMENT')
	application_record_element.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(application_record_element,'SHORT-NAME')
	short_name.text=ARDT_element_shortname

	# ARDT_element_type = [APDT, ARDT, AADT, IDT]

	if ARDT_element_type == 'APDT':

		type_tref=ET.SubElement(application_record_element,'TYPE-TREF')
		type_tref.text=f'/SharedElements/ApplicationDataTypes/Primitive/{data_type}'
		type_tref.attrib={'DEST':'APPLICATION-PRIMITIVE-DATA-TYPE'}

	elif ARDT_element_type == 'AADT':

		type_tref=ET.SubElement(application_record_element,'TYPE-TREF')
		type_tref.text=f'/SharedElements/ApplicationDataTypes/Array/{data_type}'
		type_tref.attrib={'DEST':'APPLICATION-ARRAY-DATA-TYPE'}

	elif ARDT_element_type == 'ARDT':

		type_tref=ET.SubElement(application_record_element,'TYPE-TREF')
		type_tref.text=f'/SharedElements/ApplicationDataTypes/Record/{data_type}'
		type_tref.attrib={'DEST':'APPLICATION-RECORD-DATA-TYPE'}

	elif ARDT_element_type == 'IDT':
		type_tref=ET.SubElement(application_record_element,'TYPE-TREF')
		type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{data_type}'
		type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

	else :
		print ('Invalid ARDT_element_type')

# 3. AADT

def ApplicationArrayDataType_Fixed(Array_folder_elements, ApplicationArrayDataType_Fixed_shortname, data_type, array_size): #completed

	a=processor.value_to_str(array_size)
	application_array_data_type=ET.SubElement(Array_folder_elements,'APPLICATION-ARRAY-DATA-TYPE')
	application_array_data_type.attrib={'UUID':rng.generate_uuid()} #99540e2c-05ec-4a85-94bb-9a3999ac57fe'}
	short_name=ET.SubElement(application_array_data_type,'SHORT-NAME')
	short_name.text=ApplicationArrayDataType_Fixed_shortname
	category=ET.SubElement(application_array_data_type,'CATEGORY')
	category.text='ARRAY'
	element=ET.SubElement(application_array_data_type,'ELEMENT')
	element.attrib={'UUID':rng.generate_uuid()} #7391c5fe-50b6-4b88-bc63-ec1975221a4f'}
	short_name=ET.SubElement(element,'SHORT-NAME')
	short_name.text='Element'
	category=ET.SubElement(element,'CATEGORY')
	category.text='VALUE'
	type_tref=ET.SubElement(element,'TYPE-TREF')
	type_tref.text=f'/SharedElements/ApplicationDataTypes/Primitive/{data_type}' #'/SharedElements/ApplicationDataTypes/Primitive/ApplicationPrimitiveDataType'
	type_tref.attrib={'DEST':'APPLICATION-PRIMITIVE-DATA-TYPE'}
	array_size_semantics=ET.SubElement(element,'ARRAY-SIZE-SEMANTICS')
	array_size_semantics.text='FIXED-SIZE'
	max_number_of_elements=ET.SubElement(element,'MAX-NUMBER-OF-ELEMENTS')
	max_number_of_elements.text=a

def ApplicationArrayDataType_Variable(Array_folder_elements, ApplicationArrayDataType_Variable_shortname, data_type, array_size):#completed
	a=processor.value_to_str(array_size)
	application_array_data_type=ET.SubElement(Array_folder_elements,'APPLICATION-ARRAY-DATA-TYPE')
	application_array_data_type.attrib={'UUID':rng.generate_uuid()} 
	short_name=ET.SubElement(application_array_data_type,'SHORT-NAME')
	short_name.text=ApplicationArrayDataType_Variable_shortname
	category=ET.SubElement(application_array_data_type,'CATEGORY')
	category.text='ARRAY'
	element=ET.SubElement(application_array_data_type,'ELEMENT')
	element.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(element,'SHORT-NAME')
	short_name.text='Element'
	category=ET.SubElement(element,'CATEGORY')
	category.text='VALUE'
	type_tref=ET.SubElement(element,'TYPE-TREF')
	type_tref.text=f'/SharedElements/ApplicationDataTypes/Primitive/{data_type}' #'/SharedElements/ApplicationDataTypes/Primitive/ApplicationPrimitiveDataType'
	type_tref.attrib={'DEST':'APPLICATION-PRIMITIVE-DATA-TYPE'}
	array_size_semantics=ET.SubElement(element,'ARRAY-SIZE-SEMANTICS')
	array_size_semantics.text='VARIABLE-SIZE'
	max_number_of_elements=ET.SubElement(element,'MAX-NUMBER-OF-ELEMENTS')
	max_number_of_elements.text=a


####### Compu method #########

def CompuMethod_IDENTICAL(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed

	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='IDENTICAL'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}


def CompuMethod_bitfield_text(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed
	global compu_scales
	
	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='BITFIELD_TEXTTABLE'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}
	compu_internal_to_phys=ET.SubElement(compu_method,'COMPU-INTERNAL-TO-PHYS')
	compu_scales=ET.SubElement(compu_internal_to_phys,'COMPU-SCALES')
	
def bitfield_text_compu_scale(mask_val, ll, ul, enum):#completed
	c=processor.value_to_str(mask_val)
	d=processor.value_to_str(ll)
	e=processor.value_to_str(ul)
	f=processor.value_to_str(enum)
	compu_scale=ET.SubElement(compu_scales,'COMPU-SCALE')
	mask=ET.SubElement(compu_scale,'MASK')
	mask.text=c
	lower_limit=ET.SubElement(compu_scale,'LOWER-LIMIT')
	lower_limit.text=d
	upper_limit=ET.SubElement(compu_scale,'UPPER-LIMIT')
	upper_limit.text=e
	compu_const=ET.SubElement(compu_scale,'COMPU-CONST')
	vt=ET.SubElement(compu_const,'VT')
	vt.text=f


def CompuMethod_linear(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed
	global compu_scales
	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='LINEAR'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}
	compu_internal_to_phys=ET.SubElement(compu_method,'COMPU-INTERNAL-TO-PHYS')
	compu_scales=ET.SubElement(compu_internal_to_phys,'COMPU-SCALES')
	
def linear_compu_scale(num_a,num_b,den_a):#completed
	c=processor.value_to_str(num_a)
	d=processor.value_to_str(num_b)
	e=processor.value_to_str(den_a)

	compu_scale=ET.SubElement(compu_scales,'COMPU-SCALE')
	compu_rational_coeffs=ET.SubElement(compu_scale,'COMPU-RATIONAL-COEFFS')
	compu_numerator=ET.SubElement(compu_rational_coeffs,'COMPU-NUMERATOR')
	v2=ET.SubElement(compu_numerator,'V')
	v2.text=c
	v3=ET.SubElement(compu_numerator,'V')
	v3.text=d
	compu_denominator=ET.SubElement(compu_rational_coeffs,'COMPU-DENOMINATOR')
	v4=ET.SubElement(compu_denominator,'V')
	v4.text=e


def CompuMethod_rat_func(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed
	global compu_scales
 
	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='RAT_FUNC'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}
	compu_internal_to_phys=ET.SubElement(compu_method,'COMPU-INTERNAL-TO-PHYS')
	compu_scales=ET.SubElement(compu_internal_to_phys,'COMPU-SCALES')
	
def rat_func_compu_scale(num_a,num_b,num_c,den_a,den_b,den_c):#completed
	c=processor.value_to_str(num_a)
	d=processor.value_to_str(num_b)
	e=processor.value_to_str(num_c)
	f=processor.value_to_str(den_a)
	g=processor.value_to_str(den_b)
	h=processor.value_to_str(den_c)
	
	compu_scale=ET.SubElement(compu_scales,'COMPU-SCALE')
	compu_rational_coeffs=ET.SubElement(compu_scale,'COMPU-RATIONAL-COEFFS')
	compu_numerator=ET.SubElement(compu_rational_coeffs,'COMPU-NUMERATOR')
	v5=ET.SubElement(compu_numerator,'V')
	v5.text=c
	v6=ET.SubElement(compu_numerator,'V')
	v6.text=d
	v7=ET.SubElement(compu_numerator,'V')
	v7.text=e
	compu_denominator=ET.SubElement(compu_rational_coeffs,'COMPU-DENOMINATOR')
	v8=ET.SubElement(compu_denominator,'V')
	v8.text=f
	v9=ET.SubElement(compu_denominator,'V')
	v9.text=g
	v10=ET.SubElement(compu_denominator,'V')
	v10.text=h


def CompuMethod_Scale_rat_text(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed
	global compu_scales, compu_internal_to_phys
	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	desc=ET.SubElement(compu_method,'DESC')
	l_2=ET.SubElement(desc,'L-2')
	l_2.text='S'
	l_2.attrib={'L':'FOR-ALL'}
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='SCALE_RATIONAL_AND_TEXTTABLE'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}
	compu_internal_to_phys=ET.SubElement(compu_method,'COMPU-INTERNAL-TO-PHYS')
	compu_scales=ET.SubElement(compu_internal_to_phys,'COMPU-SCALES')
	
def Scale_rat_text_compu_scale(ll,ul,num_a,num_b,num_c,den_a,den_b,den_c):#completed
	a=processor.value_to_str(ll)
	b=processor.value_to_str(ul)
	c=processor.value_to_str(num_a)
	d=processor.value_to_str(num_b)
	e=processor.value_to_str(num_c)
	f=processor.value_to_str(den_a)
	g=processor.value_to_str(den_b)
	h=processor.value_to_str(den_c)
	compu_scale=ET.SubElement(compu_scales,'COMPU-SCALE')
	lower_limit=ET.SubElement(compu_scale,'LOWER-LIMIT')
	lower_limit.text=a
	upper_limit=ET.SubElement(compu_scale,'UPPER-LIMIT')
	upper_limit.text=b
	compu_rational_coeffs=ET.SubElement(compu_scale,'COMPU-RATIONAL-COEFFS')
	compu_numerator=ET.SubElement(compu_rational_coeffs,'COMPU-NUMERATOR')
	v1=ET.SubElement(compu_numerator,'V')
	v1.text=c
	v2=ET.SubElement(compu_numerator,'V')
	v2.text=d
	v3=ET.SubElement(compu_numerator,'V')
	v3.text=e
	compu_denominator=ET.SubElement(compu_rational_coeffs,'COMPU-DENOMINATOR')
	v4=ET.SubElement(compu_denominator,'V')
	v4.text=f
	v5=ET.SubElement(compu_denominator,'V')
	v5.text=g
	v6=ET.SubElement(compu_denominator,'V')
	v6.text=h
	
	# compu_scale9=ET.SubElement(compu_scales4,'COMPU-SCALE')
	# lower_limit7=ET.SubElement(compu_scale9,'LOWER-LIMIT')
	# lower_limit7.text='16'
	# upper_limit7=ET.SubElement(compu_scale9,'UPPER-LIMIT')
	# upper_limit7.text='16'
	# compu_const6=ET.SubElement(compu_scale9,'COMPU-CONST')
	# vt6=ET.SubElement(compu_const6,'VT')
	# vt6.text='sdcd'
	
	# compu_scale10=ET.SubElement(compu_scales4,'COMPU-SCALE')
	# lower_limit8=ET.SubElement(compu_scale10,'LOWER-LIMIT')
	# lower_limit8.text='17'
	# upper_limit8=ET.SubElement(compu_scale10,'UPPER-LIMIT')
	# upper_limit8.text='17'
	# compu_const7=ET.SubElement(compu_scale10,'COMPU-CONST')
	# vt7=ET.SubElement(compu_const7,'VT')
	# vt7.text='sdcd1'
	
def Scale_rat_text_compu_default_value(cm_DefaultValue):#completed
	a=processor.value_to_str(cm_DefaultValue)
	compu_default_value=ET.SubElement(compu_internal_to_phys,'COMPU-DEFAULT-VALUE')
	v7=ET.SubElement(compu_default_value,'V')
	v7.text=a


def CompuMethod_Scale_linear_text(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed
	global compu_scales, compu_internal_to_phys

	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	desc=ET.SubElement(compu_method,'DESC')
	l_2=ET.SubElement(desc,'L-2')
	l_2.text='Scale_linear_And_texttable_CompuMethod'
	l_2.attrib={'L':'FOR-ALL'}
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='SCALE_LINEAR_AND_TEXTTABLE'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}
	compu_internal_to_phys=ET.SubElement(compu_method,'COMPU-INTERNAL-TO-PHYS')
	compu_scales=ET.SubElement(compu_internal_to_phys,'COMPU-SCALES')

def Scale_linear_text_compu_scale(ll,ul,num_a,num_b,den_a):#completed
	a=processor.value_to_str(ll)
	b=processor.value_to_str(ul)
	c=processor.value_to_str(num_a)
	d=processor.value_to_str(num_b)
	e=processor.value_to_str(den_a)
	
	compu_scale=ET.SubElement(compu_scales,'COMPU-SCALE')
	lower_limit=ET.SubElement(compu_scale,'LOWER-LIMIT')
	lower_limit.text=a
	upper_limit=ET.SubElement(compu_scale,'UPPER-LIMIT')
	upper_limit.text=b
	compu_rational_coeffs=ET.SubElement(compu_scale,'COMPU-RATIONAL-COEFFS')
	compu_numerator=ET.SubElement(compu_rational_coeffs,'COMPU-NUMERATOR')
	v=ET.SubElement(compu_numerator,'V')
	v.text=c
	v1=ET.SubElement(compu_numerator,'V')
	v1.text=d
	compu_denominator=ET.SubElement(compu_rational_coeffs,'COMPU-DENOMINATOR')
	v2=ET.SubElement(compu_denominator,'V')
	v2.text=e

def Scale_linear_text_compu_DefaultValue(cm_DefaultValue):#completed
	a=processor.value_to_str(cm_DefaultValue)
	compu_default_value=ET.SubElement(compu_internal_to_phys,'COMPU-DEFAULT-VALUE')
	v=ET.SubElement(compu_default_value,'V')
	v.text=a


def CompuMethod_tab_nointp(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed
	global compu_scales, compu_internal_to_phys
	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	desc3=ET.SubElement(compu_method,'DESC')
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='TAB_NOINTP'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}
	compu_internal_to_phys=ET.SubElement(compu_method,'COMPU-INTERNAL-TO-PHYS')
	compu_scales=ET.SubElement(compu_internal_to_phys,'COMPU-SCALES')

def tab_nointp_compu_Scale(value, enum):#completed
	a=processor.value_to_str(value)
	b=processor.value_to_str(enum)
	compu_scale=ET.SubElement(compu_scales,'COMPU-SCALE')
	# desc=ET.SubElement(compu_scale,'DESC')
	# l_2=ET.SubElement(desc,'L-2')
	# l_2.attrib={'L':'AA'}
	lower_limit=ET.SubElement(compu_scale,'LOWER-LIMIT')
	lower_limit.text=a
	upper_limit=ET.SubElement(compu_scale,'UPPER-LIMIT')
	upper_limit.text=a
	compu_const=ET.SubElement(compu_scale,'COMPU-CONST')
	v=ET.SubElement(compu_const,'V')
	v.text=b

def tab_nointp_compu_Scale_DefaultValue(cm_DefaultValue):#completed
	a=processor.value_to_str(cm_DefaultValue)
	compu_default_value=ET.SubElement(compu_internal_to_phys,'COMPU-DEFAULT-VALUE')
	vf=ET.SubElement(compu_default_value,'VF')
	vf.text=a


def CompuMethod_text(CompuMethods_shared_folder_elements, compu_method_shortname, unit):#completed
	global compu_scales, compu_internal_to_phys
	compu_method=ET.SubElement(CompuMethods_shared_folder_elements,'COMPU-METHOD')
	compu_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(compu_method,'SHORT-NAME')
	short_name.text=compu_method_shortname
	category=ET.SubElement(compu_method,'CATEGORY')
	category.text='TEXTTABLE'
	unit_ref=ET.SubElement(compu_method,'UNIT-REF')
	unit_ref.text=f'/AUTOSAR/AUTOSAR_PhysicalUnits/Units/{unit}'
	unit_ref.attrib={'DEST':'UNIT'}
	compu_internal_to_phys=ET.SubElement(compu_method,'COMPU-INTERNAL-TO-PHYS')
	compu_scales=ET.SubElement(compu_internal_to_phys,'COMPU-SCALES')

def text_compu_Scale(value,enum):#completed
	a=processor.value_to_str(value)
	b=processor.value_to_str(enum)
	
	compu_scale=ET.SubElement(compu_scales,'COMPU-SCALE')
	lower_limit=ET.SubElement(compu_scale,'LOWER-LIMIT')
	lower_limit.text=a
	upper_limit=ET.SubElement(compu_scale,'UPPER-LIMIT')
	upper_limit.text=a
	compu_const=ET.SubElement(compu_scale,'COMPU-CONST')
	vt=ET.SubElement(compu_const,'VT')
	vt.text=b

def text_compu_DefaultValue(cm_DefaultValue):#completed
	a=processor.value_to_str(cm_DefaultValue)
	compu_default_value=ET.SubElement(compu_internal_to_phys,'COMPU-DEFAULT-VALUE')
	v=ET.SubElement(compu_default_value,'V')
	v.text=a

########## other shared elements 	##############

def ConstantSpecification(ConstantSpecifications_folder_elements,constant_spec_shortname, constant_spec_Val):#completed
	a=processor.value_to_str(constant_spec_Val)
	constant_specification=ET.SubElement(ConstantSpecifications_folder_elements,'CONSTANT-SPECIFICATION')
	constant_specification.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(constant_specification,'SHORT-NAME')
	short_name.text=constant_spec_shortname
	value_spec=ET.SubElement(constant_specification,'VALUE-SPEC')
	numerical_value_specification=ET.SubElement(value_spec,'NUMERICAL-VALUE-SPECIFICATION')
	short_label=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_label.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a

def DataConstr_phy(DataConstr_folder_elements, DataConstr_shortname,ll,ul):#completed
	a=processor.value_to_str(ll)
	b=processor.value_to_str(ul)
	data_constr=ET.SubElement(DataConstr_folder_elements,'DATA-CONSTR')
	data_constr.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(data_constr,'SHORT-NAME')
	short_name.text=DataConstr_shortname
	data_constr_rules=ET.SubElement(data_constr,'DATA-CONSTR-RULES')
	data_constr_rule=ET.SubElement(data_constr_rules,'DATA-CONSTR-RULE')
	phys_constrs=ET.SubElement(data_constr_rule,'PHYS-CONSTRS')
	lower_limit=ET.SubElement(phys_constrs,'LOWER-LIMIT')
	lower_limit.text=a
	upper_limit=ET.SubElement(phys_constrs,'UPPER-LIMIT')
	upper_limit.text=b

def DataConstr_Int(DataConstr_folder_elements, DataConstr_shortname,ll,ul):#completed
	a=processor.value_to_str(ll)
	b=processor.value_to_str(ul)
	data_constr=ET.SubElement(DataConstr_folder_elements,'DATA-CONSTR')
	data_constr.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(data_constr,'SHORT-NAME')
	short_name.text=DataConstr_shortname
	data_constr_rules=ET.SubElement(data_constr,'DATA-CONSTR-RULES')
	data_constr_rule=ET.SubElement(data_constr_rules,'DATA-CONSTR-RULE')
	phys_constrs=ET.SubElement(data_constr_rule,'INTERNAL-CONSTRS')
	lower_limit=ET.SubElement(phys_constrs,'LOWER-LIMIT')
	lower_limit.text=a
	upper_limit=ET.SubElement(phys_constrs,'UPPER-LIMIT')
	upper_limit.text=b

def SwcImplementation(SwcImplementation_folder_elements,SwcImplementation_shortname,SWC_IB):#completed

	swc_implementation=ET.SubElement(SwcImplementation_folder_elements,'SWC-IMPLEMENTATION')
	swc_implementation.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(swc_implementation,'SHORT-NAME')
	short_name.text=SwcImplementation_shortname
	programming_language=ET.SubElement(swc_implementation,'PROGRAMMING-LANGUAGE')
	programming_language.text='C'
	resource_consumption=ET.SubElement(swc_implementation,'RESOURCE-CONSUMPTION')
	resource_consumption.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(resource_consumption,'SHORT-NAME')
	short_name.text='ResourceConsumption'
	sw_version=ET.SubElement(swc_implementation,'SW-VERSION')
	sw_version.text='1.0.0.0'
	behavior_ref=ET.SubElement(swc_implementation,'BEHAVIOR-REF')
	behavior_ref.text=f'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/{SWC_IB}' #run if else for each SWC type
	behavior_ref.attrib={'DEST':'SWC-INTERNAL-BEHAVIOR'}

def SwAddrMethod(SwAddrMethod_folder_elements, SwAddrMethod_shortname, mem_alloc_policy, mem_section_type):#completed
	sw_addr_method=ET.SubElement(SwAddrMethod_folder_elements,'SW-ADDR-METHOD')
	sw_addr_method.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(sw_addr_method,'SHORT-NAME')
	short_name.text=SwAddrMethod_shortname
	memory_allocation_keyword_policy=ET.SubElement(sw_addr_method,'MEMORY-ALLOCATION-KEYWORD-POLICY')
	memory_allocation_keyword_policy.text= mem_alloc_policy  #'ADDR-METHOD-SHORT-NAME', 'ADDR-METHOD-SHORT-NAME-AND-ALIGNMENT'
	section_type=ET.SubElement(sw_addr_method,'SECTION-TYPE')
	section_type.text= mem_section_type # 'CODE', 'CALIBRATION-VARIABLES', VAR, CONST, CALPARM, CONFIG-DATA, EXCLUDE-FROM-Flash

########## DATA type mapping set ##########

def DataTypeMappingSet(DataTypemappingSets_folder_elements, CurrentSWC_shortname): #completed
	global data_type_maps

	data_type_mapping_set=ET.SubElement(DataTypemappingSets_folder_elements,'DATA-TYPE-MAPPING-SET')
	data_type_mapping_set.attrib={'UUID': rng.generate_uuid()}
	short_name=ET.SubElement(data_type_mapping_set,'SHORT-NAME')
	short_name.text=f'DTMS_{CurrentSWC_shortname}'
	data_type_maps=ET.SubElement(data_type_mapping_set,'DATA-TYPE-MAPS')

def data_type_map(adt,idt): #completed
	data_type_map=ET.SubElement(data_type_maps,'DATA-TYPE-MAP')
	application_data_type_ref=ET.SubElement(data_type_map,'APPLICATION-DATA-TYPE-REF')
	application_data_type_ref.text=f'/SharedElements/ApplicationDataTypes/Array/{adt}'
	application_data_type_ref.attrib={'DEST':'APPLICATION-ARRAY-DATA-TYPE'}
	implementation_data_type_ref=ET.SubElement(data_type_map,'IMPLEMENTATION-DATA-TYPE-REF')
	implementation_data_type_ref.text=f'/SharedElements/ImplementationDataTypes/{idt}'
	implementation_data_type_ref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

########## Implementation Data type ########## ApplicationArrayDataType_Fixed

def ImplementationDataType_ArrayFixed(ImplementationDataTypes_folder_elements, IDT_shortname, arraysize_fixed, IDT):#completed
	a=processor.value_to_str(arraysize_fixed)
	implementation_data_type=ET.SubElement(ImplementationDataTypes_folder_elements,'IMPLEMENTATION-DATA-TYPE')
	implementation_data_type.attrib={'UUID': rng.generate_uuid()}
	short_name=ET.SubElement(implementation_data_type,'SHORT-NAME')
	short_name.text=IDT_shortname
	category=ET.SubElement(implementation_data_type,'CATEGORY')
	category.text='ARRAY'
	sub_elements=ET.SubElement(implementation_data_type,'SUB-ELEMENTS')
	implementation_data_type_element=ET.SubElement(sub_elements,'IMPLEMENTATION-DATA-TYPE-ELEMENT')
	implementation_data_type_element.attrib={'UUID': rng.generate_uuid()}
	short_name=ET.SubElement(implementation_data_type_element,'SHORT-NAME')
	short_name.text='SubElement'
	category=ET.SubElement(implementation_data_type_element,'CATEGORY')
	category.text='TYPE_REFERENCE'
	array_size=ET.SubElement(implementation_data_type_element,'ARRAY-SIZE')
	array_size.text=a
	array_size_semantics=ET.SubElement(implementation_data_type_element,'ARRAY-SIZE-SEMANTICS')
	array_size_semantics.text='FIXED-SIZE'
	sw_data_def_props=ET.SubElement(implementation_data_type_element,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	implementation_data_type_ref=ET.SubElement(sw_data_def_props_conditional,'IMPLEMENTATION-DATA-TYPE-REF')
	implementation_data_type_ref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{IDT}'
	implementation_data_type_ref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

def ImplementationDataType_ArrayVariable(ImplementationDataTypes_folder_elements, IDT_shortname, arraysize_variable, IDT):#completed but need to revisit after actual implementation
	a=processor.value_to_str(arraysize_variable)
	implementation_data_type=ET.SubElement(ImplementationDataTypes_folder_elements,'IMPLEMENTATION-DATA-TYPE')
	implementation_data_type.attrib={'UUID': rng.generate_uuid()}
	short_name=ET.SubElement(implementation_data_type,'SHORT-NAME')
	short_name.text=IDT_shortname
	category=ET.SubElement(implementation_data_type,'CATEGORY')
	category.text='ARRAY'
	sub_elements=ET.SubElement(implementation_data_type,'SUB-ELEMENTS')
	implementation_data_type_element=ET.SubElement(sub_elements,'IMPLEMENTATION-DATA-TYPE-ELEMENT')
	implementation_data_type_element.attrib={'UUID': rng.generate_uuid()}
	short_name=ET.SubElement(implementation_data_type_element,'SHORT-NAME')
	short_name.text='SubElement'
	category=ET.SubElement(implementation_data_type_element,'CATEGORY')
	category.text='TYPE_REFERENCE'
	array_size=ET.SubElement(implementation_data_type_element,'ARRAY-SIZE')
	array_size.text=a
	array_size_semantics=ET.SubElement(implementation_data_type_element,'ARRAY-SIZE-SEMANTICS')
	array_size_semantics.text='VARIABLE-SIZE'
	sw_data_def_props=ET.SubElement(implementation_data_type_element,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	implementation_data_type_ref=ET.SubElement(sw_data_def_props_conditional,'IMPLEMENTATION-DATA-TYPE-REF')
	implementation_data_type_ref.text=f'/SharedElements/ImplementationDataTypes/{IDT}'
	implementation_data_type_ref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

		# make a structure element with following structure 
		# <IMPLEMENTATION-DATA-TYPE UUID="53ec3bfc-5a92-4d42-b31b-8e29e99a2b46">
        #       <SHORT-NAME>STRUCTURE_ImplementationDataType1</SHORT-NAME>
        #       <CATEGORY>STRUCTURE</CATEGORY>
        #       <SUB-ELEMENTS>
        #         <IMPLEMENTATION-DATA-TYPE-ELEMENT UUID="31f01782-3ce8-4dbe-81d1-0d5fb89bef99">
        #           <SHORT-NAME>SubElement</SHORT-NAME>
        #           <CATEGORY>TYPE_REFERENCE</CATEGORY>
        #           <SW-DATA-DEF-PROPS>
        #             <SW-DATA-DEF-PROPS-VARIANTS>
        #               <SW-DATA-DEF-PROPS-CONDITIONAL>
        #                 <IMPLEMENTATION-DATA-TYPE-REF DEST="IMPLEMENTATION-DATA-TYPE">/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/uint16</IMPLEMENTATION-DATA-TYPE-REF>
        #               </SW-DATA-DEF-PROPS-CONDITIONAL>
        #             </SW-DATA-DEF-PROPS-VARIANTS>
        #           </SW-DATA-DEF-PROPS>
        #         </IMPLEMENTATION-DATA-TYPE-ELEMENT>
        #         <IMPLEMENTATION-DATA-TYPE-ELEMENT UUID="83bd06cb-a4ff-4d55-bd3d-1a691b582d46">
        #           <SHORT-NAME>SubElement1</SHORT-NAME>
        #           <CATEGORY>TYPE_REFERENCE</CATEGORY>
        #           <SW-DATA-DEF-PROPS>
        #             <SW-DATA-DEF-PROPS-VARIANTS>
        #               <SW-DATA-DEF-PROPS-CONDITIONAL>
        #                 <IMPLEMENTATION-DATA-TYPE-REF DEST="IMPLEMENTATION-DATA-TYPE">/SharedElements/ImplementationDataTypes/ARRAY_ImplementationDataType</IMPLEMENTATION-DATA-TYPE-REF>
        #               </SW-DATA-DEF-PROPS-CONDITIONAL>
        #             </SW-DATA-DEF-PROPS-VARIANTS>
        #           </SW-DATA-DEF-PROPS>
        #         </IMPLEMENTATION-DATA-TYPE-ELEMENT>
        #       </SUB-ELEMENTS>
        #     </IMPLEMENTATION-DATA-TYPE>

		#   <IMPLEMENTATION-DATA-TYPE UUID="21f9a013-317d-4a6a-8c1d-cdc72f7df8f5">
		#   <SHORT-NAME>ARRAY_ImplementationDataType</SHORT-NAME>
		#   <CATEGORY>ARRAY</CATEGORY>
		#   <SW-DATA-DEF-PROPS>
		#     <SW-DATA-DEF-PROPS-VARIANTS>
		#       <SW-DATA-DEF-PROPS-CONDITIONAL />
		#     </SW-DATA-DEF-PROPS-VARIANTS>
		#   </SW-DATA-DEF-PROPS>
		#   <SUB-ELEMENTS>
		#     <IMPLEMENTATION-DATA-TYPE-ELEMENT UUID="5512b8b7-a43f-436f-bb18-47a903ad1e17">
		#       <SHORT-NAME>SubElement</SHORT-NAME>
		#       <CATEGORY>TYPE_REFERENCE</CATEGORY>
		#       <ARRAY-SIZE>15</ARRAY-SIZE>
		#       <ARRAY-SIZE-SEMANTICS>FIXED-SIZE</ARRAY-SIZE-SEMANTICS>
		#       <SW-DATA-DEF-PROPS>
		#         <SW-DATA-DEF-PROPS-VARIANTS>
		#           <SW-DATA-DEF-PROPS-CONDITIONAL>
		#             <IMPLEMENTATION-DATA-TYPE-REF DEST="IMPLEMENTATION-DATA-TYPE">/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/uint16</IMPLEMENTATION-DATA-TYPE-REF>
		#           </SW-DATA-DEF-PROPS-CONDITIONAL>
		#         </SW-DATA-DEF-PROPS-VARIANTS>
		#       </SW-DATA-DEF-PROPS>
		#     </IMPLEMENTATION-DATA-TYPE-ELEMENT>
		#   </SUB-ELEMENTS>
		# </IMPLEMENTATION-DATA-TYPE>



		# for dtms

		# <APPLICATION-ARRAY-DATA-TYPE UUID="d5f3c7e9-dd94-4d37-888e-b6e44b01cc5a">
        #           <SHORT-NAME>ApplicationArrayDataType_Variable</SHORT-NAME>
        #           <CATEGORY>ARRAY</CATEGORY>
        #           <ELEMENT UUID="fef3f4b8-d9bd-4cb1-94b8-4403e665c4fa">
        #             <SHORT-NAME>Element</SHORT-NAME>
        #             <CATEGORY>VALUE</CATEGORY>
        #             <TYPE-TREF DEST="APPLICATION-PRIMITIVE-DATA-TYPE">/SharedElements/ApplicationDataTypes/Primitive/ApplicationPrimitiveDataType</TYPE-TREF>
        #             <ARRAY-SIZE-SEMANTICS>VARIABLE-SIZE</ARRAY-SIZE-SEMANTICS>
        #             <MAX-NUMBER-OF-ELEMENTS>15</MAX-NUMBER-OF-ELEMENTS>
        #           </ELEMENT>
        #         </APPLICATION-ARRAY-DATA-TYPE>

		#         <IMPLEMENTATION-DATA-TYPE UUID="ccd15817-26a8-424d-8c87-3f3d70b5ee9d">
		#   <SHORT-NAME>Struct_Array_ImplementationDataType</SHORT-NAME>
		#   <CATEGORY>STRUCTURE</CATEGORY>
		#   <SUB-ELEMENTS>
		#     <IMPLEMENTATION-DATA-TYPE-ELEMENT UUID="3f61bc0d-d829-4ab0-9e22-7de6a25972e3">
		#       <SHORT-NAME>SubElement1</SHORT-NAME>
		#       <CATEGORY>TYPE_REFERENCE</CATEGORY>
		#       <SW-DATA-DEF-PROPS>
		#         <SW-DATA-DEF-PROPS-VARIANTS>
		#           <SW-DATA-DEF-PROPS-CONDITIONAL>
		#             <IMPLEMENTATION-DATA-TYPE-REF DEST="IMPLEMENTATION-DATA-TYPE">/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/uint8</IMPLEMENTATION-DATA-TYPE-REF>
		#           </SW-DATA-DEF-PROPS-CONDITIONAL>
		#         </SW-DATA-DEF-PROPS-VARIANTS>
		#       </SW-DATA-DEF-PROPS>
		#     </IMPLEMENTATION-DATA-TYPE-ELEMENT>
		#     <IMPLEMENTATION-DATA-TYPE-ELEMENT UUID="dc530c9c-3b65-4707-99c3-842e2d2b7788">
		#       <SHORT-NAME>SubElement</SHORT-NAME>
		#       <CATEGORY>ARRAY</CATEGORY>
		#       <SUB-ELEMENTS>
		#         <IMPLEMENTATION-DATA-TYPE-ELEMENT UUID="af21d788-9aea-4789-b7d0-8665f2d0c8c7">
		#           <SHORT-NAME>SubElement</SHORT-NAME>
		#           <CATEGORY>TYPE_REFERENCE</CATEGORY>
		#           <ARRAY-SIZE>15</ARRAY-SIZE>
		#           <ARRAY-SIZE-SEMANTICS>VARIABLE-SIZE</ARRAY-SIZE-SEMANTICS>
		#           <SW-DATA-DEF-PROPS>
		#             <SW-DATA-DEF-PROPS-VARIANTS>
		#               <SW-DATA-DEF-PROPS-CONDITIONAL>
		#                 <IMPLEMENTATION-DATA-TYPE-REF DEST="IMPLEMENTATION-DATA-TYPE">/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/uint16</IMPLEMENTATION-DATA-TYPE-REF>
		#               </SW-DATA-DEF-PROPS-CONDITIONAL>
		#             </SW-DATA-DEF-PROPS-VARIANTS>
		#           </SW-DATA-DEF-PROPS>
		#         </IMPLEMENTATION-DATA-TYPE-ELEMENT>
		#       </SUB-ELEMENTS>
		#       <SW-DATA-DEF-PROPS>
		#         <SW-DATA-DEF-PROPS-VARIANTS>
		#           <SW-DATA-DEF-PROPS-CONDITIONAL />
		#         </SW-DATA-DEF-PROPS-VARIANTS>
		#       </SW-DATA-DEF-PROPS>
		#     </IMPLEMENTATION-DATA-TYPE-ELEMENT>
		#   </SUB-ELEMENTS>
		# </IMPLEMENTATION-DATA-TYPE>

		# def Struct_Array_ImplementationDataType():

		# 	implementation_data_type5=ET.SubElement(elements10,'IMPLEMENTATION-DATA-TYPE')
		# 	implementation_data_type5.attrib={'UUID':rng.generate_uuid()} #ccd15817-26a8-424d-8c87-3f3d70b5ee9d'}
		# 	short_name46=ET.SubElement(implementation_data_type5,'SHORT-NAME')
		# 	short_name46.text='Struct_Array_ImplementationDataType'
		# 	category27=ET.SubElement(implementation_data_type5,'CATEGORY')
		# 	category27.text='STRUCTURE'
		# 	sub_elements4=ET.SubElement(implementation_data_type5,'SUB-ELEMENTS')
		# 	implementation_data_type_element5=ET.SubElement(sub_elements4,'IMPLEMENTATION-DATA-TYPE-ELEMENT')
		# 	implementation_data_type_element5.attrib={'UUID':rng.generate_uuid()} #3f61bc0d-d829-4ab0-9e22-7de6a25972e3'}
		# 	short_name47=ET.SubElement(implementation_data_type_element5,'SHORT-NAME')
		# 	short_name47.text='SubElement1'
		# 	category28=ET.SubElement(implementation_data_type_element5,'CATEGORY')
		# 	category28.text='TYPE_REFERENCE'
		# 	sw_data_def_props11=ET.SubElement(implementation_data_type_element5,'SW-DATA-DEF-PROPS')
		# 	sw_data_def_props_variants11=ET.SubElement(sw_data_def_props11,'SW-DATA-DEF-PROPS-VARIANTS')
		# 	sw_data_def_props_conditional11=ET.SubElement(sw_data_def_props_variants11,'SW-DATA-DEF-PROPS-CONDITIONAL')
		# 	implementation_data_type_ref9=ET.SubElement(sw_data_def_props_conditional11,'IMPLEMENTATION-DATA-TYPE-REF')
		# 	implementation_data_type_ref9.text='/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/uint8'
		# 	implementation_data_type_ref9.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
		# 	implementation_data_type_element6=ET.SubElement(sub_elements4,'IMPLEMENTATION-DATA-TYPE-ELEMENT')
		# 	implementation_data_type_element6.attrib={'UUID':rng.generate_uuid()} #dc530c9c-3b65-4707-99c3-842e2d2b7788'}
		# 	short_name48=ET.SubElement(implementation_data_type_element6,'SHORT-NAME')
		# 	short_name48.text='SubElement'
		# 	category29=ET.SubElement(implementation_data_type_element6,'CATEGORY')
		# 	category29.text='ARRAY'
		# 	sub_elements5=ET.SubElement(implementation_data_type_element6,'SUB-ELEMENTS')
		# 	implementation_data_type_element7=ET.SubElement(sub_elements5,'IMPLEMENTATION-DATA-TYPE-ELEMENT')
		# 	implementation_data_type_element7.attrib={'UUID':rng.generate_uuid()} #af21d788-9aea-4789-b7d0-8665f2d0c8c7'}
		# 	short_name49=ET.SubElement(implementation_data_type_element7,'SHORT-NAME')
		# 	short_name49.text='SubElement'
		# 	category30=ET.SubElement(implementation_data_type_element7,'CATEGORY')
		# 	category30.text='TYPE_REFERENCE'
		# 	array_size3=ET.SubElement(implementation_data_type_element7,'ARRAY-SIZE')
		# 	array_size3.text='15'
		# 	array_size_semantics6=ET.SubElement(implementation_data_type_element7,'ARRAY-SIZE-SEMANTICS')
		# 	array_size_semantics6.text='VARIABLE-SIZE'
		# 	sw_data_def_props12=ET.SubElement(implementation_data_type_element7,'SW-DATA-DEF-PROPS')
		# 	sw_data_def_props_variants12=ET.SubElement(sw_data_def_props12,'SW-DATA-DEF-PROPS-VARIANTS')
		# 	sw_data_def_props_conditional12=ET.SubElement(sw_data_def_props_variants12,'SW-DATA-DEF-PROPS-CONDITIONAL')
		# 	implementation_data_type_ref10=ET.SubElement(sw_data_def_props_conditional12,'IMPLEMENTATION-DATA-TYPE-REF')
		# 	implementation_data_type_ref10.text='/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/uint16'
		# 	implementation_data_type_ref10.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
		# 	sw_data_def_props13=ET.SubElement(implementation_data_type_element6,'SW-DATA-DEF-PROPS')
		# 	sw_data_def_props_variants13=ET.SubElement(sw_data_def_props13,'SW-DATA-DEF-PROPS-VARIANTS')
		# 	sw_data_def_props_conditional13=ET.SubElement(sw_data_def_props_variants13,'SW-DATA-DEF-PROPS-CONDITIONAL')
		# 	implementation_data_type6=ET.SubElement(elements10,'IMPLEMENTATION-DATA-TYPE')
		# 	implementation_data_type6.attrib={'UUID':rng.generate_uuid()} #79fa9e8f-a805-43da-b4b5-ac42d2a23ff0'}
		# 	short_name50=ET.SubElement(implementation_data_type6,'SHORT-NAME')
		# 	short_name50.text='TypeTref_ImplementationDataType'
		# 	category31=ET.SubElement(implementation_data_type6,'CATEGORY')
		# 	category31.text='TYPE_REFERENCE'
		# 	sw_data_def_props14=ET.SubElement(implementation_data_type6,'SW-DATA-DEF-PROPS')
		# 	sw_data_def_props_variants14=ET.SubElement(sw_data_def_props14,'SW-DATA-DEF-PROPS-VARIANTS')
		# 	sw_data_def_props_conditional14=ET.SubElement(sw_data_def_props_variants14,'SW-DATA-DEF-PROPS-CONDITIONAL')
		# 	implementation_data_type_ref11=ET.SubElement(sw_data_def_props_conditional14,'IMPLEMENTATION-DATA-TYPE-REF')
		# 	implementation_data_type_ref11.text='/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/float32'
		# 	implementation_data_type_ref11.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

def ImplementationDataType(ImplementationDataTypes_folder_elements, IDT_shortname, IDT):#completed

	implementation_data_type=ET.SubElement(ImplementationDataTypes_folder_elements,'IMPLEMENTATION-DATA-TYPE')
	implementation_data_type.attrib={'UUID': rng.generate_uuid()}
	short_name=ET.SubElement(implementation_data_type,'SHORT-NAME')
	short_name.text=IDT_shortname
	category=ET.SubElement(implementation_data_type,'CATEGORY')
	category.text='VALUE'
	sw_data_def_props=ET.SubElement(implementation_data_type,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	base_type_ref=ET.SubElement(sw_data_def_props_conditional,'BASE-TYPE-REF')
	base_type_ref.text=f'/AUTOSAR/AUTOSAR_Platform/BaseTypes/{IDT}'
	base_type_ref.attrib={'DEST':'SW-BASE-TYPE'}

def ImplementationDataType_Structure(ImplementationDataTypes_folder_elements, IDT_shortname):#completed
	global IDT_elements
	implementation_data_type=ET.SubElement(ImplementationDataTypes_folder_elements,'IMPLEMENTATION-DATA-TYPE')
	implementation_data_type.attrib={'UUID': rng.generate_uuid()}
	short_name=ET.SubElement(implementation_data_type,'SHORT-NAME')
	short_name.text=IDT_shortname
	category=ET.SubElement(implementation_data_type,'CATEGORY')
	category.text='STRUCTURE'
	IDT_elements=ET.SubElement(implementation_data_type,'SUB-ELEMENTS')

def ImplementationDataType_Record_elements(IDT_element_shortname, IDT):#completed 
    implementation_data_type_element=ET.SubElement(IDT_elements,'IMPLEMENTATION-DATA-TYPE-ELEMENT')
    implementation_data_type_element.attrib={'UUID': rng.generate_uuid()}
    short_name=ET.SubElement(implementation_data_type_element,'SHORT-NAME')
    short_name.text=IDT_element_shortname
    category=ET.SubElement(implementation_data_type_element,'CATEGORY')
    category.text='TYPE_REFERENCE'
    sw_data_def_props=ET.SubElement(implementation_data_type_element,'SW-DATA-DEF-PROPS')
    sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
    sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
    implementation_data_type_ref=ET.SubElement(sw_data_def_props_conditional,'IMPLEMENTATION-DATA-TYPE-REF')
    implementation_data_type_ref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{IDT}'
    implementation_data_type_ref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

############ Interfaces ##################### 

def ClientServerInterface(ClientServer_folder_elements, IF_Name):#completed
    global client_server_interface
    client_server_interface=ET.SubElement(ClientServer_folder_elements,'CLIENT-SERVER-INTERFACE')
    client_server_interface.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(client_server_interface,'SHORT-NAME')
    short_name.text= IF_Name
    is_service=ET.SubElement(client_server_interface,'IS-SERVICE')
    is_service.text='false'

def ClientServerInterface_Opr():#completed
	global operations
	operations=ET.SubElement(client_server_interface,'OPERATIONS')

def ClientServerInterface_CSOpr(Operation_shortname):#completed
	global client_server_operation
	client_server_operation=ET.SubElement(operations,'CLIENT-SERVER-OPERATION')
	client_server_operation.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(client_server_operation,'SHORT-NAME')
	short_name.text=Operation_shortname

def ClientServerInterface_Args():#completed
	global arguments
	arguments=ET.SubElement(client_server_operation,'ARGUMENTS')

def ClientServerInterface_Arg(Argument_shortname, type_tref_adt):#completed

    argument_data_prototype=ET.SubElement(arguments,'ARGUMENT-DATA-PROTOTYPE')
    argument_data_prototype.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(argument_data_prototype,'SHORT-NAME')
    short_name.text=Argument_shortname
    sw_data_def_props=ET.SubElement(argument_data_prototype,'SW-DATA-DEF-PROPS')
    sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
    sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
    sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
    sw_impl_policy.text='STANDARD'
    type_tref=ET.SubElement(argument_data_prototype,'TYPE-TREF')
    type_tref.text=f'/SharedElements/ApplicationDataTypes/Primitive/{type_tref_adt}'
    type_tref.attrib={'DEST':'APPLICATION-PRIMITIVE-DATA-TYPE'}
    direction=ET.SubElement(argument_data_prototype,'DIRECTION')
    direction.text='IN'
    server_argument_impl_policy=ET.SubElement(argument_data_prototype,'SERVER-ARGUMENT-IMPL-POLICY')
    server_argument_impl_policy.text='USE-ARGUMENT-TYPE'


def ModeDeclarationGroup(ModeSwitch_folder_elements, ModeDeclarationGroup_shortname, mode_Category,Init_Mode):#completed
	global mode_declarations
	mode_declaration_group=ET.SubElement(ModeSwitch_folder_elements,'MODE-DECLARATION-GROUP')
	mode_declaration_group.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(mode_declaration_group,'SHORT-NAME')
	short_name.text=ModeDeclarationGroup_shortname
	category=ET.SubElement(mode_declaration_group,'CATEGORY')
	category.text= mode_Category
	initial_mode_ref=ET.SubElement(mode_declaration_group,'INITIAL-MODE-REF')
	initial_mode_ref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{short_name.text}/{Init_Mode}'
	initial_mode_ref.attrib={'DEST':'MODE-DECLARATION'}
	mode_declarations=ET.SubElement(mode_declaration_group,'MODE-DECLARATIONS')
    
def ModeDeclarationGroup_Exp(ModeDeclaration_shortname):#completed

    mode_declaration=ET.SubElement(mode_declarations,'MODE-DECLARATION')
    mode_declaration.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(mode_declaration,'SHORT-NAME')
    short_name.text=ModeDeclaration_shortname

def ModeSwitchInterface(ModeSwitch_folder_elements, ModeSwitchInterface_shortname, ModeDeclarationGroup_shortname):#completed

	mode_switch_interface=ET.SubElement(ModeSwitch_folder_elements,'MODE-SWITCH-INTERFACE')
	mode_switch_interface.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(mode_switch_interface,'SHORT-NAME')
	short_name.text=ModeSwitchInterface_shortname
	is_service=ET.SubElement(mode_switch_interface,'IS-SERVICE')
	is_service.text='false'
	mode_group=ET.SubElement(mode_switch_interface,'MODE-GROUP')
	mode_group.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(mode_group,'SHORT-NAME')
	short_name.text='ModeGroup'
	type_tref=ET.SubElement(mode_group,'TYPE-TREF')
	type_tref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{ModeDeclarationGroup_shortname}'
	type_tref.attrib={'DEST':'MODE-DECLARATION-GROUP'}


def NvDataInterface(NvData_folder_elements, IF_Name):#completed
 
    global nv_data_interface


    nv_data_interface=ET.SubElement(NvData_folder_elements,'NV-DATA-INTERFACE')
    nv_data_interface.attrib={'UUID':rng.generate_uuid()} #8a4989b3-88e2-4e47-b98f-591e75c76b17'}
    short_name=ET.SubElement(nv_data_interface,'SHORT-NAME')
    short_name.text=IF_Name
    is_service=ET.SubElement(nv_data_interface,'IS-SERVICE')
    is_service.text='false'

def NvDataInterface_DE():#completed
	global nv_datas
	nv_datas=ET.SubElement(nv_data_interface,'NV-DATAS')

def NvDataInterface_VDP(nv_datas_shortname, type_tref_adt):#completed
    variable_data_prototype=ET.SubElement(nv_datas,'VARIABLE-DATA-PROTOTYPE')
    variable_data_prototype.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_data_prototype,'SHORT-NAME')
    short_name.text=nv_datas_shortname
    sw_data_def_props=ET.SubElement(variable_data_prototype,'SW-DATA-DEF-PROPS')
    sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
    sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
    sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
    sw_calibration_access.text='READ-WRITE'
    sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
    sw_impl_policy.text='STANDARD'
    type_tref=ET.SubElement(variable_data_prototype,'TYPE-TREF')
    type_tref.text=f'/SharedElements/ApplicationDataTypes/Primitive/{type_tref_adt}'
    type_tref.attrib={'DEST':'APPLICATION-PRIMITIVE-DATA-TYPE'}


def ParameterInterface(Parameter_folder_elements, IF_Name):#completed
	global parameter_interface
	parameter_interface=ET.SubElement(Parameter_folder_elements,'PARAMETER-INTERFACE')
	parameter_interface.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_interface,'SHORT-NAME')
	short_name.text= IF_Name
	is_service=ET.SubElement(parameter_interface,'IS-SERVICE')
	is_service.text='false'

def ParameterInterface_DE():#completed
	global parameters
	parameters=ET.SubElement(parameter_interface,'PARAMETERS')

def ParameterInterface_VDP(Parameter_shortname, type_tref_adt):#completed

	parameter_data_prototype=ET.SubElement(parameters,'PARAMETER-DATA-PROTOTYPE')
	parameter_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_data_prototype,'SHORT-NAME')
	short_name.text=Parameter_shortname
	sw_data_def_props=ET.SubElement(parameter_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text='READ-WRITE'
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text='STANDARD'
	type_tref=ET.SubElement(parameter_data_prototype,'TYPE-TREF')
	type_tref.text=f'/SharedElements/ApplicationDataTypes/Primitive/{type_tref_adt}'
	type_tref.attrib={'DEST':'APPLICATION-PRIMITIVE-DATA-TYPE'}


def SenderReceiverInterface(SenderReceiver_folder_elements, IF_Name):#completed
	global sender_receiver_interface
	sender_receiver_interface=ET.SubElement(SenderReceiver_folder_elements,'SENDER-RECEIVER-INTERFACE')
	sender_receiver_interface.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(sender_receiver_interface,'SHORT-NAME')
	short_name.text= IF_Name
	is_service=ET.SubElement(sender_receiver_interface,'IS-SERVICE')
	is_service.text='false'

def SenderReceiverInterface_DE():#completed
	global data_elements
	data_elements=ET.SubElement(sender_receiver_interface,'DATA-ELEMENTS')

def SenderReceiverInterface_VDP(DataElement_shortname, type_tref_adt):#completed

	variable_data_prototype=ET.SubElement(data_elements,'VARIABLE-DATA-PROTOTYPE')
	variable_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_data_prototype,'SHORT-NAME')
	short_name.text=DataElement_shortname
	sw_data_def_props=ET.SubElement(variable_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text='READ-WRITE'
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text='STANDARD'
	type_tref=ET.SubElement(variable_data_prototype,'TYPE-TREF')
	type_tref.text=f'/SharedElements/ApplicationDataTypes/Primitive/{type_tref_adt}'
	type_tref.attrib={'DEST':'APPLICATION-PRIMITIVE-DATA-TYPE'}
	
	# variable_data_prototype14=ET.SubElement(data_elements6,'VARIABLE-DATA-PROTOTYPE')
	# variable_data_prototype14.attrib={'UUID':rng.generate_uuid()} #6862a5ea-8794-4906-9f54-50624e9d6044'}
	# short_name117=ET.SubElement(variable_data_prototype14,'SHORT-NAME')
	# short_name117.text='DataElement1'
	# sw_data_def_props40=ET.SubElement(variable_data_prototype14,'SW-DATA-DEF-PROPS')
	# sw_data_def_props_variants40=ET.SubElement(sw_data_def_props40,'SW-DATA-DEF-PROPS-VARIANTS')
	# sw_data_def_props_conditional40=ET.SubElement(sw_data_def_props_variants40,'SW-DATA-DEF-PROPS-CONDITIONAL')
	# sw_impl_policy26=ET.SubElement(sw_data_def_props_conditional40,'SW-IMPL-POLICY')
	# sw_impl_policy26.text='STANDARD'
	# type_tref32=ET.SubElement(variable_data_prototype14,'TYPE-TREF')
	# type_tref32.text='/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/uint32'
	# type_tref32.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
	# # invalidation_policys1=ET.SubElement(sender_receiver_interface,'INVALIDATION-POLICYS')
	# # invalidation_policy1=ET.SubElement(invalidation_policys1,'INVALIDATION-POLICY')
	# # data_element_ref1=ET.SubElement(invalidation_policy1,'DATA-ELEMENT-REF')
	# # data_element_ref1.text='/SharedElements/PortInterfaces/SenderReceiver/SenderReceiverInterface/DataElement'
	# # data_element_ref1.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}
	# # handle_invalid1=ET.SubElement(invalidation_policy1,'HANDLE-INVALID')
	# # handle_invalid1.text='KEEP'


def TriggerInterface(Trigger_folder_elements, IF_Name):#completed
	global trigger_interface
	trigger_interface=ET.SubElement(Trigger_folder_elements,'TRIGGER-INTERFACE')
	trigger_interface.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(trigger_interface,'SHORT-NAME')
	short_name.text= IF_Name
	is_service=ET.SubElement(trigger_interface,'IS-SERVICE')
	is_service.text='false'

def TriggerInterface_trigs():#completed
	global triggers
	triggers=ET.SubElement(trigger_interface,'TRIGGERS')

def TriggerInterface_trig(trigger_shortname, cse_code, cse_code_factor):#completed
	a=processor.value_to_str(cse_code)
	b=processor.value_to_str(cse_code_factor)	
	trigger=ET.SubElement(triggers,'TRIGGER')
	trigger.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(trigger,'SHORT-NAME')
	short_name.text=trigger_shortname
	trigger_period=ET.SubElement(trigger,'TRIGGER-PERIOD')
	cse_code=ET.SubElement(trigger_period,'CSE-CODE')
	cse_code.text=a
	cse_code_factor=ET.SubElement(trigger_period,'CSE-CODE-FACTOR')
	cse_code_factor.text=b

########## PORTS ###########

# comspec of each port type (basically ports are categorised based on which type of interface they referenced) is different that is why we need to create each port type separately
# and this prt of the port , we will do it later

def create_ports(swc_type): #completed
	global ports


	if swc_type == 'ApplicationSwComponentType':
		
		ports=ET.SubElement(application_sw_component_type,'PORTS')

	elif swc_type == 'ComplexDeviceDriverSwComponentType':

		ports=ET.SubElement(complex_device_driver_sw_component_type,'PORTS')

	elif swc_type == 'EcuAbstractionSwComponentType':

		ports=ET.SubElement(ecu_abstraction_sw_component_type,'PORTS')

	elif swc_type == 'NvBlockSwComponentType':		

		ports=ET.SubElement(nv_block_sw_component_type,'PORTS')

	elif swc_type == 'SensorActuatorSwComponentType':		

		ports=ET.SubElement(sensor_actuator_sw_component_type,'PORTS')

	elif swc_type == 'ServiceProxySwComponentType':		

		ports=ET.SubElement(service_proxy_sw_component_type,'PORTS')

	elif swc_type == 'ServiceSwComponentType':		

		ports=ET.SubElement(service_sw_component_type,'PORTS')

	elif swc_type == 'ParameterSwComponentType':		

		ports=ET.SubElement(parameter_sw_component_type,'PORTS')	
		

	
	else :
		pass



	# 'ApplicationSwComponentType': my_application_function,  # Change the function name here
    #     'ComplexDeviceDriverSwComponentType': my_complex_device_driver_function,  # Change the function name here
    #     'EcuAbstractionSwComponentType': my_ecu_abstraction_function,  # Change the function name here
    #     'NvBlockSwComponentType': my_nv_block_function,  # Change the function name here
    #     'ParameterSwComponentType': my_parameter_function,  # Change the function name here
    #     'SensorActuatorSwComponentType': my_sensor_actuator_function,  # Change the function name here
    #     'ServiceProxySwComponentType': my_service_proxy_function,  # Change the function name here
    #     'ServiceSwComponentType': my_service_function  # Change the function name here


def RPort_SR(Port_shortname, referred_IF): 	#partially completed
	
    r_port_prototype=ET.SubElement(ports,'R-PORT-PROTOTYPE')
    r_port_prototype.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(r_port_prototype,'SHORT-NAME')
    short_name.text= Port_shortname
    required_interface_tref=ET.SubElement(r_port_prototype,'REQUIRED-INTERFACE-TREF')
    required_interface_tref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{referred_IF}'
    required_interface_tref.attrib={'DEST':'SENDER-RECEIVER-INTERFACE'} 
 
def RPort_CS(Port_shortname, referred_IF):	#partially completed
 
	r_port_prototype=ET.SubElement(ports,'R-PORT-PROTOTYPE')
	r_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(r_port_prototype,'SHORT-NAME')
	short_name.text=Port_shortname
	required_interface_tref=ET.SubElement(r_port_prototype,'REQUIRED-INTERFACE-TREF')
	required_interface_tref.text=f'/SharedElements/PortInterfaces/ClientServer/{referred_IF}'
	required_interface_tref.attrib={'DEST':'CLIENT-SERVER-INTERFACE'}
 
def RPort_msi(Port_shortname, referred_IF): #partially completed
 
	r_port_prototype=ET.SubElement(ports,'R-PORT-PROTOTYPE')
	r_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(r_port_prototype,'SHORT-NAME')
	short_name.text=Port_shortname
	required_interface_tref=ET.SubElement(r_port_prototype,'REQUIRED-INTERFACE-TREF')
	required_interface_tref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{referred_IF}'
	required_interface_tref.attrib={'DEST':'MODE-SWITCH-INTERFACE'}

def RPort_nvd(Port_shortname, referred_IF): #partially completed
 
	r_port_prototype=ET.SubElement(ports,'R-PORT-PROTOTYPE')
	r_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(r_port_prototype,'SHORT-NAME')
	short_name.text=Port_shortname
	required_interface_tref=ET.SubElement(r_port_prototype,'REQUIRED-INTERFACE-TREF')
	required_interface_tref.text=f'/SharedElements/PortInterfaces/NvData/{referred_IF}'
	required_interface_tref.attrib={'DEST':'NV-DATA-INTERFACE'}
 
def RPort_prm(Port_shortname, referred_IF): #partially completed
 
	r_port_prototype=ET.SubElement(ports,'R-PORT-PROTOTYPE')
	r_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(r_port_prototype,'SHORT-NAME')
	short_name.text=Port_shortname
	required_interface_tref=ET.SubElement(r_port_prototype,'REQUIRED-INTERFACE-TREF')
	required_interface_tref.text=f'/SharedElements/PortInterfaces/Parameter/{referred_IF}'
	required_interface_tref.attrib={'DEST':'PARAMETER-INTERFACE'}

def RPort_trigger(Port_shortname, referred_IF): #partially completed

	r_port_prototype=ET.SubElement(ports,'R-PORT-PROTOTYPE')
	r_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(r_port_prototype,'SHORT-NAME')
	short_name.text=Port_shortname
	required_interface_tref=ET.SubElement(r_port_prototype,'REQUIRED-INTERFACE-TREF')
	required_interface_tref.text=f'/SharedElements/PortInterfaces/Trigger/{referred_IF}'
	required_interface_tref.attrib={'DEST':'TRIGGER-INTERFACE'}
	
def PPort_SR(Port_shortname, referred_IF): #partially completed
     
	p_port_prototype=ET.SubElement(ports,'P-PORT-PROTOTYPE')
	p_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(p_port_prototype,'SHORT-NAME')
	short_name.text= Port_shortname
	provided_interface_tref=ET.SubElement(p_port_prototype,'PROVIDED-INTERFACE-TREF')
	provided_interface_tref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{referred_IF}'
	provided_interface_tref.attrib={'DEST':'SENDER-RECEIVER-INTERFACE'} 
 
def PPort_CS(Port_shortname, referred_IF): #partially completed
 
	p_port_prototype=ET.SubElement(ports,'P-PORT-PROTOTYPE')
	p_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(p_port_prototype,'SHORT-NAME')
	short_name.text= Port_shortname
	provided_interface_tref=ET.SubElement(p_port_prototype,'PROVIDED-INTERFACE-TREF')
	provided_interface_tref.text=f'/SharedElements/PortInterfaces/ClientServer/{referred_IF}'
	provided_interface_tref.attrib={'DEST':'CLIENT-SERVER-INTERFACE'}

def PPort_msi(Port_shortname, referred_IF): #partially completed
 
	p_port_prototype=ET.SubElement(ports,'P-PORT-PROTOTYPE')
	p_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(p_port_prototype,'SHORT-NAME')
	short_name.text= Port_shortname
	provided_interface_tref=ET.SubElement(p_port_prototype,'PROVIDED-INTERFACE-TREF')
	provided_interface_tref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{referred_IF}'
	provided_interface_tref.attrib={'DEST':'MODE-SWITCH-INTERFACE'}

def PPort_nvd(Port_shortname, referred_IF): #partially completed

	p_port_prototype=ET.SubElement(ports,'P-PORT-PROTOTYPE')
	p_port_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(p_port_prototype,'SHORT-NAME')
	short_name.text= Port_shortname
	provided_interface_tref=ET.SubElement(p_port_prototype,'PROVIDED-INTERFACE-TREF')
	provided_interface_tref.text=f'/SharedElements/PortInterfaces/NvData/{referred_IF}'
	provided_interface_tref.attrib={'DEST':'NV-DATA-INTERFACE'} 

########## IB ###########

def internal_behaviors(CurrentInternalBehaviors_shortname,swc_type): #partially completed if other type of component to be added
    
	global swc_internal_behavior, IB_shortname

	if swc_type == 'ApplicationSwComponentType':

		internal_behaviors=ET.SubElement(application_sw_component_type,'INTERNAL-BEHAVIORS')

	elif swc_type == 'ComplexDeviceDriverSwComponentType':

		internal_behaviors=ET.SubElement(complex_device_driver_sw_component_type,'INTERNAL-BEHAVIORS')

	elif swc_type == 'EcuAbstractionSwComponentType':

		internal_behaviors=ET.SubElement(ecu_abstraction_sw_component_type,'INTERNAL-BEHAVIORS')

	elif swc_type == 'NvBlockSwComponentType':		

		internal_behaviors=ET.SubElement(nv_block_sw_component_type,'INTERNAL-BEHAVIORS')

	elif swc_type == 'SensorActuatorSwComponentType':		

		internal_behaviors=ET.SubElement(sensor_actuator_sw_component_type,'INTERNAL-BEHAVIORS')

	elif swc_type == 'ServiceProxySwComponentType':		

		internal_behaviors=ET.SubElement(service_proxy_sw_component_type,'INTERNAL-BEHAVIORS')

	elif swc_type == 'ServiceSwComponentType':		

		internal_behaviors=ET.SubElement(service_sw_component_type,'INTERNAL-BEHAVIORS')						
	
	else :
		pass

	
	swc_internal_behavior=ET.SubElement(internal_behaviors,'SWC-INTERNAL-BEHAVIOR')
	swc_internal_behavior.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(swc_internal_behavior,'SHORT-NAME')
	short_name.text=CurrentInternalBehaviors_shortname
	IB_shortname = CurrentInternalBehaviors_shortname


def ConstantMemory():#completed
	global constant_memorys
	constant_memorys=ET.SubElement(swc_internal_behavior,'CONSTANT-MEMORYS')

def ConstantMemory_PDP(ConstantMemory_shortname, type_tref_adt, Init_val, sw_calibration_access_Enum, sw_impl_policy_Enum): #completed
	a=processor.value_to_str(Init_val)
	parameter_data_prototype=ET.SubElement(constant_memorys,'PARAMETER-DATA-PROTOTYPE')
	parameter_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_data_prototype,'SHORT-NAME')
	short_name.text=ConstantMemory_shortname
	sw_data_def_props=ET.SubElement(parameter_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text=sw_calibration_access_Enum
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text=sw_impl_policy_Enum
	type_tref=ET.SubElement(parameter_data_prototype,'TYPE-TREF')
	type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{type_tref_adt}'
	type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
	init_value=ET.SubElement(parameter_data_prototype,'INIT-VALUE')
	numerical_value_specification=ET.SubElement(init_value,'NUMERICAL-VALUE-SPECIFICATION')
	short_labe=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_labe.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a


def DataTYPEMAPPINGREFS():#completed
	global data_type_mapping_refs
 
	data_type_mapping_refs=ET.SubElement(swc_internal_behavior,'DATA-TYPE-MAPPING-REFS')

def DataTYPEMAPPINGREF(CurrentSWC_shortname):#completed
 
	data_type_mapping_ref=ET.SubElement(data_type_mapping_refs,'DATA-TYPE-MAPPING-REF')
	data_type_mapping_ref.text=f'/SharedElements/DataTypemappingSets/DTMS_{CurrentSWC_shortname}'
	data_type_mapping_ref.attrib={'DEST':'DATA-TYPE-MAPPING-SET'}


def StaticMemory():#completed
	global static_memorys
	static_memorys=ET.SubElement(swc_internal_behavior,'STATIC-MEMORYS')

def StaticMemory_VDP(StaticMemory_shortname, type_tref_adt, Init_val, sw_calibration_access_Enum, sw_impl_policy_Enum):#completed
	a=processor.value_to_str(Init_val)

	variable_data_prototype=ET.SubElement(static_memorys,'VARIABLE-DATA-PROTOTYPE')
	variable_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_data_prototype,'SHORT-NAME')
	short_name.text=StaticMemory_shortname
	sw_data_def_props=ET.SubElement(variable_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text=sw_calibration_access_Enum
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text=sw_impl_policy_Enum
	type_tref=ET.SubElement(variable_data_prototype,'TYPE-TREF')
	type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{type_tref_adt}'
	type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

	init_value=ET.SubElement(variable_data_prototype,'INIT-VALUE')
	numerical_value_specification=ET.SubElement(init_value,'NUMERICAL-VALUE-SPECIFICATION')
	short_labe=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_labe.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a

	#constant reference need to do something for init value


	# init_value2=ET.SubElement(variable_data_prototype,'INIT-VALUE')
	# constant_reference1=ET.SubElement(init_value2,'CONSTANT-REFERENCE')
	# short_label5=ET.SubElement(constant_reference1,'SHORT-LABEL')
	# short_label5.text='ReferenceToConstant'
	# constant_ref1=ET.SubElement(constant_reference1,'CONSTANT-REF')
	# constant_ref1.text='/SharedElements/ConstantSpecifications/ApplicationSwComponentType_StaticMemory'
	# constant_ref1.attrib={'DEST':'CONSTANT-SPECIFICATION'}


def ArTypedPerInstanceMemory():#completed
	global ar_typed_per_instance_memorys
 
	ar_typed_per_instance_memorys=ET.SubElement(swc_internal_behavior,'AR-TYPED-PER-INSTANCE-MEMORYS')

def ArTypedPerInstanceMemory_VDP(ArTypedPerInstanceMemory_shortname, type_tref_adt, Init_val, sw_calibration_access_Enum, sw_impl_policy_Enum):	#completed
	a=processor.value_to_str(Init_val)
	variable_data_prototype=ET.SubElement(ar_typed_per_instance_memorys,'VARIABLE-DATA-PROTOTYPE')
	variable_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_data_prototype,'SHORT-NAME')
	short_name.text=ArTypedPerInstanceMemory_shortname
	sw_data_def_props=ET.SubElement(variable_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text=sw_calibration_access_Enum
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text=sw_impl_policy_Enum
	type_tref=ET.SubElement(variable_data_prototype,'TYPE-TREF')
	type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{type_tref_adt}'
	type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}

	init_value=ET.SubElement(variable_data_prototype,'INIT-VALUE')
	numerical_value_specification=ET.SubElement(init_value,'NUMERICAL-VALUE-SPECIFICATION')
	short_labe=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_labe.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a


def ExplicitInterRunnableVariable():#completed
	global explicit_inter_runnable_variables
 
	explicit_inter_runnable_variables=ET.SubElement(swc_internal_behavior,'EXPLICIT-INTER-RUNNABLE-VARIABLES')

def ExplicitInterRunnableVariable_VDP(ExplicitInterRunnableVariable_shortname, type_tref_adt, Init_val, sw_calibration_access_Enum, sw_impl_policy_Enum):#completed
	a=processor.value_to_str(Init_val)
	variable_data_prototype=ET.SubElement(explicit_inter_runnable_variables,'VARIABLE-DATA-PROTOTYPE')
	variable_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_data_prototype,'SHORT-NAME')
	short_name.text=ExplicitInterRunnableVariable_shortname
	sw_data_def_props=ET.SubElement(variable_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text=sw_calibration_access_Enum
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text= sw_impl_policy_Enum
	type_tref=ET.SubElement(variable_data_prototype,'TYPE-TREF')
	type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{type_tref_adt}'
	type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
	
	init_value=ET.SubElement(variable_data_prototype,'INIT-VALUE')
	numerical_value_specification=ET.SubElement(init_value,'NUMERICAL-VALUE-SPECIFICATION')
	short_labe=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_labe.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a


def handle_termination_and_restart(handle_termination_and_restart_Enum):#completed
 
	handle_termination_and_restart=ET.SubElement(swc_internal_behavior,'HANDLE-TERMINATION-AND-RESTART')
	handle_termination_and_restart.text=handle_termination_and_restart_Enum


def ImplicitInterRunnableVariable():#completed
	global implicit_inter_runnable_variables
 
	implicit_inter_runnable_variables=ET.SubElement(swc_internal_behavior,'IMPLICIT-INTER-RUNNABLE-VARIABLES')

def ImplicitInterRunnableVariable_VDP(implicit_inter_runnable_variable_shortname, type_tref_adt, Init_val, sw_calibration_access_Enum, sw_impl_policy_Enum):#completed
	a=processor.value_to_str(Init_val)
	variable_data_prototype=ET.SubElement(implicit_inter_runnable_variables,'VARIABLE-DATA-PROTOTYPE')
	variable_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_data_prototype,'SHORT-NAME')
	short_name.text=implicit_inter_runnable_variable_shortname
	sw_data_def_props=ET.SubElement(variable_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text=sw_calibration_access_Enum
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text= sw_impl_policy_Enum
	type_tref=ET.SubElement(variable_data_prototype,'TYPE-TREF')
	type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{type_tref_adt}'
	type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
	
	init_value=ET.SubElement(variable_data_prototype,'INIT-VALUE')
	numerical_value_specification=ET.SubElement(init_value,'NUMERICAL-VALUE-SPECIFICATION')
	short_labe=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_labe.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a


def PerInstanceParameter():#completed
	global per_instance_parameters
	per_instance_parameters=ET.SubElement(swc_internal_behavior,'PER-INSTANCE-PARAMETERS')

def PerInstanceParameter_PDP(per_instance_parameters_shortname, type_tref_adt, Init_val, sw_calibration_access_Enum, sw_impl_policy_Enum):#completed

	a=processor.value_to_str(Init_val)
	parameter_data_prototype=ET.SubElement(per_instance_parameters,'PARAMETER-DATA-PROTOTYPE')
	parameter_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_data_prototype,'SHORT-NAME')
	short_name.text=per_instance_parameters_shortname
	sw_data_def_props=ET.SubElement(parameter_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text=sw_calibration_access_Enum
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text=sw_impl_policy_Enum
	type_tref=ET.SubElement(parameter_data_prototype,'TYPE-TREF')
	type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{type_tref_adt}'
	type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
	init_value=ET.SubElement(parameter_data_prototype,'INIT-VALUE')
	numerical_value_specification=ET.SubElement(init_value,'NUMERICAL-VALUE-SPECIFICATION')
	short_labe=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_labe.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a


def SharedParameter(): #completed
    
    global shared_parameters
    shared_parameters=ET.SubElement(swc_internal_behavior,'SHARED-PARAMETERS')

def SharedParameter_PDP(SharedParameter_shortname, type_tref_adt, Init_val, sw_calibration_access_Enum, sw_impl_policy_Enum):#completed
	a=processor.value_to_str(Init_val)
	parameter_data_prototype=ET.SubElement(shared_parameters,'PARAMETER-DATA-PROTOTYPE')
	parameter_data_prototype.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_data_prototype,'SHORT-NAME')
	short_name.text=SharedParameter_shortname
	sw_data_def_props=ET.SubElement(parameter_data_prototype,'SW-DATA-DEF-PROPS')
	sw_data_def_props_variants=ET.SubElement(sw_data_def_props,'SW-DATA-DEF-PROPS-VARIANTS')
	sw_data_def_props_conditional=ET.SubElement(sw_data_def_props_variants,'SW-DATA-DEF-PROPS-CONDITIONAL')
	sw_calibration_access=ET.SubElement(sw_data_def_props_conditional,'SW-CALIBRATION-ACCESS')
	sw_calibration_access.text=sw_calibration_access_Enum
	sw_impl_policy=ET.SubElement(sw_data_def_props_conditional,'SW-IMPL-POLICY')
	sw_impl_policy.text=sw_impl_policy_Enum
	type_tref=ET.SubElement(parameter_data_prototype,'TYPE-TREF')
	type_tref.text=f'/AUTOSAR/AUTOSAR_Platform/ImplementationDataTypes/{type_tref_adt}'
	type_tref.attrib={'DEST':'IMPLEMENTATION-DATA-TYPE'}
	init_value=ET.SubElement(parameter_data_prototype,'INIT-VALUE')
	numerical_value_specification=ET.SubElement(init_value,'NUMERICAL-VALUE-SPECIFICATION')
	short_labe=ET.SubElement(numerical_value_specification,'SHORT-LABEL')
	short_labe.text='Value'
	value=ET.SubElement(numerical_value_specification,'VALUE')
	value.text=a


def supports_multiple_instantiation(supports_multiple_instantiation_enum):#completed
    supports_multiple_instantiation=ET.SubElement(swc_internal_behavior,'SUPPORTS-MULTIPLE-INSTANTIATION')
    supports_multiple_instantiation.text=supports_multiple_instantiation_enum

########## RTE Events ###########

def RTE_Event():#completed
	global Rte_events
	Rte_events=ET.SubElement(swc_internal_behavior,'EVENTS')

def AsynchronousServerCallReturnsEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname):#completed 

    asynchronous_server_call_returns_event=ET.SubElement(Rte_events,'ASYNCHRONOUS-SERVER-CALL-RETURNS-EVENT')
    asynchronous_server_call_returns_event.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(asynchronous_server_call_returns_event,'SHORT-NAME')
    short_name.text=RTE_Event_name
    start_on_event_ref=ET.SubElement(asynchronous_server_call_returns_event,'START-ON-EVENT-REF')
    start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}'
    start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
    event_source_ref=ET.SubElement(asynchronous_server_call_returns_event,'EVENT-SOURCE-REF')
    event_source_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}/{ASCP_short_name}'
    event_source_ref.attrib={'DEST':'ASYNCHRONOUS-SERVER-CALL-RESULT-POINT'}

def BackgroundEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname):#completed

	background_event=ET.SubElement(Rte_events,'BACKGROUND-EVENT')
	background_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(background_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(background_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}

def DataReceiveErrorEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed

	data_receive_error_event=ET.SubElement(Rte_events,'DATA-RECEIVE-ERROR-EVENT')
	data_receive_error_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(data_receive_error_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(data_receive_error_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #Runnable2'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	data_iref=ET.SubElement(data_receive_error_event,'DATA-IREF')
	context_r_port_ref=ET.SubElement(data_iref,'CONTEXT-R-PORT-REF')
	context_r_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/RPort_SR'
	context_r_port_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	target_data_element_ref=ET.SubElement(data_iref,'TARGET-DATA-ELEMENT-REF')
	target_data_element_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}' #'/SharedElements/PortInterfaces/SenderReceiver/SenderReceiverInterface/DataElement'
	target_data_element_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def DataReceivedEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed
 
	data_received_event=ET.SubElement(Rte_events,'DATA-RECEIVED-EVENT')
	data_received_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(data_received_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(data_received_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #Runnable3'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	data_iref=ET.SubElement(data_received_event,'DATA-IREF')
	context_r_port_ref=ET.SubElement(data_iref,'CONTEXT-R-PORT-REF')
	context_r_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}' #/SwComponentTypes/ApplSWC/ApplicationSwComponentType/RPort_SR'
	context_r_port_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	target_data_element_ref=ET.SubElement(data_iref,'TARGET-DATA-ELEMENT-REF')
	target_data_element_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}' #'/SharedElements/PortInterfaces/SenderReceiver/SenderReceiverInterface/DataElement1'
	target_data_element_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def DataSendCompletedEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, pport, DE):#completed
 
	data_send_completed_event=ET.SubElement(Rte_events,'DATA-SEND-COMPLETED-EVENT')
	data_send_completed_event.attrib={'UUID':rng.generate_uuid()} 
	short_name=ET.SubElement(data_send_completed_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(data_send_completed_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable4'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	event_source_ref=ET.SubElement(data_send_completed_event,'EVENT-SOURCE-REF')
	event_source_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}/DSP_{pport}_{DE}' #/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable4/DSP_PPort_SR_DataElement'
	event_source_ref.attrib={'DEST':'VARIABLE-ACCESS'}

def DataWriteCompletedEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, pport, DE):#completed
 
	data_write_completed_event=ET.SubElement(Rte_events,'DATA-WRITE-COMPLETED-EVENT')
	data_write_completed_event.attrib={'UUID':rng.generate_uuid()} 
	short_name=ET.SubElement(data_write_completed_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(data_write_completed_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable5'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	event_source_ref=ET.SubElement(data_write_completed_event,'EVENT-SOURCE-REF')
	event_source_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}/DWA_{pport}_{DE}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable5/DWA_PPort_SR_DataElement1'
	event_source_ref.attrib={'DEST':'VARIABLE-ACCESS'}

def ExternalTriggerOccurredEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, rport, If_name, trigger):#completed
 
	external_trigger_occurred_event=ET.SubElement(Rte_events,'EXTERNAL-TRIGGER-OCCURRED-EVENT')
	external_trigger_occurred_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(external_trigger_occurred_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(external_trigger_occurred_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable6'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	trigger_iref=ET.SubElement(external_trigger_occurred_event,'TRIGGER-IREF')
	context_r_port_ref=ET.SubElement(trigger_iref,'CONTEXT-R-PORT-REF')
	context_r_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/RPort_trigger'
	context_r_port_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	target_trigger_ref=ET.SubElement(trigger_iref,'TARGET-TRIGGER-REF')
	target_trigger_ref.text=f'/SharedElements/PortInterfaces/Trigger/{If_name}/{trigger}' #'/SharedElements/PortInterfaces/Trigger/TriggerInterface/Trigger'
	target_trigger_ref.attrib={'DEST':'TRIGGER'}

def ModeSwitchedAckEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, pport, modegroup):#completed
 
	mode_switched_ack_event=ET.SubElement(Rte_events,'MODE-SWITCHED-ACK-EVENT')
	mode_switched_ack_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(mode_switched_ack_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(mode_switched_ack_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable9'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	event_source_ref=ET.SubElement(mode_switched_ack_event,'EVENT-SOURCE-REF')
	event_source_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}/MSP_{pport}_{modegroup}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable9/MSP_PPort_msi_ModeGroup'
	event_source_ref.attrib={'DEST':'MODE-SWITCH-POINT'}

def OperationInvokedEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, pport, If_name, operation):#completed
 
	operation_invoked_event=ET.SubElement(Rte_events,'OPERATION-INVOKED-EVENT')
	operation_invoked_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(operation_invoked_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(operation_invoked_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable10'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	operation_iref=ET.SubElement(operation_invoked_event,'OPERATION-IREF')
	context_p_port_ref=ET.SubElement(operation_iref,'CONTEXT-P-PORT-REF')
	context_p_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/PPort_CS'
	context_p_port_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
	target_provided_operation_ref=ET.SubElement(operation_iref,'TARGET-PROVIDED-OPERATION-REF')
	target_provided_operation_ref.text=f'/SharedElements/PortInterfaces/ClientServer/{If_name}/{operation}' #'/SharedElements/PortInterfaces/ClientServer/ClientServerInterface/Operation1'
	target_provided_operation_ref.attrib={'DEST':'CLIENT-SERVER-OPERATION'}

def SwcModeSwitchEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, rport, If_name, modegroup, mode):#partially completed

	swc_mode_switch_event=ET.SubElement(Rte_events,'SWC-MODE-SWITCH-EVENT')
	swc_mode_switch_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(swc_mode_switch_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(swc_mode_switch_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable12'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	activation=ET.SubElement(swc_mode_switch_event,'ACTIVATION')
	activation.text='ON-TRANSITION' #other erason remaining like 'ON-ENTRY' or 'ON-EXIT'
	mode_irefs=ET.SubElement(swc_mode_switch_event,'MODE-IREFS')
	mode_iref1=ET.SubElement(mode_irefs,'MODE-IREF')
	context_port_ref=ET.SubElement(mode_iref1,'CONTEXT-PORT-REF')
	context_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/RPort_msi'
	context_port_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	context_mode_declaration_group_prototype_ref=ET.SubElement(mode_iref1,'CONTEXT-MODE-DECLARATION-GROUP-PROTOTYPE-REF')
	context_mode_declaration_group_prototype_ref.text= f'/SharedElements/PortInterfaces/ModeSwitch/{If_name}/{modegroup}' #'/SharedElements/PortInterfaces/ModeSwitch/ModeSwitchInterface/ModeGroup'
	context_mode_declaration_group_prototype_ref.attrib={'DEST':'MODE-DECLARATION-GROUP-PROTOTYPE'}
	target_mode_declaration_ref=ET.SubElement(mode_iref1,'TARGET-MODE-DECLARATION-REF')
	target_mode_declaration_ref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{modegroup}/{mode}' #'/SharedElements/PortInterfaces/ModeSwitch/ModeDeclarationGroup/ModeDeclaration2'
	target_mode_declaration_ref.attrib={'DEST':'MODE-DECLARATION'}
	mode_iref2=ET.SubElement(mode_irefs,'MODE-IREF')
	context_port_ref=ET.SubElement(mode_iref2,'CONTEXT-PORT-REF')
	context_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/RPort_msi'
	context_port_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	context_mode_declaration_group_prototype_ref=ET.SubElement(mode_iref2,'CONTEXT-MODE-DECLARATION-GROUP-PROTOTYPE-REF')
	context_mode_declaration_group_prototype_ref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{If_name}/{modegroup}' #'/SharedElements/PortInterfaces/ModeSwitch/ModeSwitchInterface/ModeGroup'
	context_mode_declaration_group_prototype_ref.attrib={'DEST':'MODE-DECLARATION-GROUP-PROTOTYPE'}
	target_mode_declaration_ref=ET.SubElement(mode_iref2,'TARGET-MODE-DECLARATION-REF')
	target_mode_declaration_ref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{modegroup}/{mode}' #'/SharedElements/PortInterfaces/ModeSwitch/ModeDeclarationGroup/ModeDeclaration'
	target_mode_declaration_ref.attrib={'DEST':'MODE-DECLARATION'}

def TimingEvent(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname, periodictime):#completed
	a=processor.value_to_str(periodictime)
	timing_event=ET.SubElement(Rte_events,'TIMING-EVENT')
	timing_event.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(timing_event,'SHORT-NAME')
	short_name.text=RTE_Event_name
	start_on_event_ref=ET.SubElement(timing_event,'START-ON-EVENT-REF')
	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable13'
	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	period=ET.SubElement(timing_event,'PERIOD')
	period.text= a

	# def TimingEvent1(RTE_Event_name,Rnbl_shortname,currentfolder, CurrentSWC_shortname):#completed
	
	# 	timing_event=ET.SubElement(Rte_events,'TIMING-EVENT')
	# 	timing_event.attrib={'UUID':rng.generate_uuid()}
	# 	short_name=ET.SubElement(timing_event,'SHORT-NAME')
	# 	short_name.text=RTE_Event_name
	# 	start_on_event_ref=ET.SubElement(timing_event,'START-ON-EVENT-REF')
	# 	start_on_event_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/Runnable15'
	# 	start_on_event_ref.attrib={'DEST':'RUNNABLE-ENTITY'}
	# 	period=ET.SubElement(timing_event,'PERIOD')
	# 	period.text='0.01'

 
########## Runnable ###########

def create_Runnable():#completed
    
	global runnables
	runnables=ET.SubElement(swc_internal_behavior,'RUNNABLES')

def Runnable_ASCRE(Rnbl_shortname,currentfolder, CurrentSWC_shortname, rport, If_name, operation):#completed
	global ASCP_short_name, runnable_entity

	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	# sw_addr_method_ref=ET.SubElement(runnable_entity,'SW-ADDR-METHOD-REF')
	# sw_addr_method_ref.text=f'/SharedElements/SwAddrMethods/{SwAddrMethod}'
	# sw_addr_method_ref.attrib={'DEST':'SW-ADDR-METHOD'}
	asynchronous_server_call_result_points=ET.SubElement(runnable_entity,'ASYNCHRONOUS-SERVER-CALL-RESULT-POINTS')
	asynchronous_server_call_result_point=ET.SubElement(asynchronous_server_call_result_points,'ASYNCHRONOUS-SERVER-CALL-RESULT-POINT')
	asynchronous_server_call_result_point.attrib={'UUID':rng.generate_uuid()}
	ASCP_short_name=ET.SubElement(asynchronous_server_call_result_point,'SHORT-NAME')
	ASCP_short_name.text='AsynchronousServerCallResultPoint'
	asynchronous_server_call_point_ref=ET.SubElement(asynchronous_server_call_result_point,'ASYNCHRONOUS-SERVER-CALL-POINT-REF')
	asynchronous_server_call_point_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{Rnbl_shortname}/ASCP_{rport}_{operation}'
	# currentfolder = ApplSWC, CurrentSWC_shortname = ApplicationSwComponentType
	asynchronous_server_call_point_ref.attrib={'DEST':'ASYNCHRONOUS-SERVER-CALL-POINT'}
	can_be_invoked_concurrently1=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently1.text='false'
	server_call_points=ET.SubElement(runnable_entity,'SERVER-CALL-POINTS')
	asynchronous_server_call_point=ET.SubElement(server_call_points,'ASYNCHRONOUS-SERVER-CALL-POINT')
	asynchronous_server_call_point.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(asynchronous_server_call_point,'SHORT-NAME')
	short_name.text=f'ASCP_{rport}_{operation}'
	operation_iref=ET.SubElement(asynchronous_server_call_point,'OPERATION-IREF')
	context_r_port_ref=ET.SubElement(operation_iref,'CONTEXT-R-PORT-REF')
	context_r_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
	context_r_port_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	target_required_operation_ref=ET.SubElement(operation_iref,'TARGET-REQUIRED-OPERATION-REF')
	target_required_operation_ref.text=f'/SharedElements/PortInterfaces/ClientServer/{If_name}/{operation}'
	target_required_operation_ref.attrib={'DEST':'CLIENT-SERVER-OPERATION'}
	timeout=ET.SubElement(asynchronous_server_call_point,'TIMEOUT')
	timeout.text='0'
	

def Runnable_BE(Rnbl_shortname):#completed
	global runnable_entity
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()} 
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	
 
def Runnable_DREE(Rnbl_shortname):#completed
	global runnable_entity 
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	

def Runnable_DRE(Rnbl_shortname):#completed
	global runnable_entity 
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	

def Runnable_DSCE(Rnbl_shortname,currentfolder, CurrentSWC_shortname, pport, If_name, DE):#completed
	global runnable_entity 
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	data_send_points=ET.SubElement(runnable_entity,'DATA-SEND-POINTS')
	variable_access=ET.SubElement(data_send_points,'VARIABLE-ACCESS')
	variable_access.attrib={'UUID':rng.generate_uuid()}
	short_name1=ET.SubElement(variable_access,'SHORT-NAME')
	short_name1.text=f'DSP_{pport}_{DE}' #'DSP_PPort_SR_DataElement'
	accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
	autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
	port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
	port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/PPort_SR'
	port_prototype_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
	target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
	target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}' #'/SharedElements/PortInterfaces/SenderReceiver/SenderReceiverInterface/DataElement'
	target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}
	

def Runnable_DWCE(Rnbl_shortname,currentfolder, CurrentSWC_shortname, pport, If_name, DE):#completed
	global runnable_entity
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	data_write_accesss=ET.SubElement(runnable_entity,'DATA-WRITE-ACCESSS')
	variable_access=ET.SubElement(data_write_accesss,'VARIABLE-ACCESS')
	variable_access.attrib={'UUID':rng.generate_uuid()}
	short_name1=ET.SubElement(variable_access,'SHORT-NAME')
	short_name1.text=f'DWA_{pport}_{DE}' #'DWA_PPort_SR_DataElement1'
	accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
	autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
	port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
	port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/PPort_SR'
	port_prototype_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
	target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
	target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}' #'/SharedElements/PortInterfaces/SenderReceiver/SenderReceiverInterface/DataElement1'
	target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}
	

def Runnable_ETOE(Rnbl_shortname):#completed
	global runnable_entity 
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	

def Runnable_MSAE(Rnbl_shortname,currentfolder, CurrentSWC_shortname, pport, If_name, modegroup):#completed
	global runnable_entity
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	mode_switch_points=ET.SubElement(runnable_entity,'MODE-SWITCH-POINTS')
	mode_switch_point=ET.SubElement(mode_switch_points,'MODE-SWITCH-POINT')
	mode_switch_point.attrib={'UUID':rng.generate_uuid()}
	short_name1=ET.SubElement(mode_switch_point,'SHORT-NAME')
	short_name1.text=f'MSP_{pport}_{modegroup}' #'MSP_PPort_msi_ModeGroup'
	mode_group_iref=ET.SubElement(mode_switch_point,'MODE-GROUP-IREF')
	context_p_port_ref=ET.SubElement(mode_group_iref,'CONTEXT-P-PORT-REF')
	context_p_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/PPort_msi'
	context_p_port_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
	target_mode_group_ref=ET.SubElement(mode_group_iref,'TARGET-MODE-GROUP-REF')
	target_mode_group_ref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{If_name}/{modegroup}' #'/SharedElements/PortInterfaces/ModeSwitch/ModeSwitchInterface/ModeGroup'
	target_mode_group_ref.attrib={'DEST':'MODE-DECLARATION-GROUP-PROTOTYPE'}
	

def Runnable_OIE(Rnbl_shortname):#completed
	global runnable_entity
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	

def Runnable_SMSE(Rnbl_shortname): #completed
	global runnable_entity
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'
	

def Runnable_TE(Rnbl_shortname):#completed
	global runnable_entity    
	runnable_entity=ET.SubElement(runnables,'RUNNABLE-ENTITY')
	runnable_entity.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(runnable_entity,'SHORT-NAME')
	short_name.text=Rnbl_shortname
	minimum_start_interval=ET.SubElement(runnable_entity,'MINIMUM-START-INTERVAL')
	minimum_start_interval.text='0'
	can_be_invoked_concurrently=ET.SubElement(runnable_entity,'CAN-BE-INVOKED-CONCURRENTLY')
	can_be_invoked_concurrently.text='false'

def Rnblsymbol(Rnbl_symbol):
	symbol=ET.SubElement(runnable_entity,'SYMBOL')
	symbol.text=Rnbl_symbol


########## Runnable access ###########



#-------------------------SR and NvD interfaces------------------------------------#


# data read access and data write access >> Implicit 

def dra ():#completed
	global data_read_accesss

	data_read_accesss=ET.SubElement(runnable_entity,'DATA-READ-ACCESSS')

def DRA_RPort_SR_DataElement(currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed
      
    variable_access=ET.SubElement(data_read_accesss,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DRA_{rport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
    port_prototype_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def DRA_RPort_nvd_NvData(currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed
 
    variable_access=ET.SubElement(data_read_accesss,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DRA_{rport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
    port_prototype_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/NvData/{If_name}/{DE}'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}


def dwa ():#completed
	global data_write_accesss
	data_write_accesss=ET.SubElement(runnable_entity,'DATA-WRITE-ACCESSS')

def DWA_PPort_SR_DataElement(currentfolder, CurrentSWC_shortname, pport, If_name, DE):#completed

    variable_access=ET.SubElement(data_write_accesss,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DWA_{pport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}'
    port_prototype_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def DWA_PPort_nvd_NvData(currentfolder, CurrentSWC_shortname, pport, If_name, DE):#completed 
 
    variable_access=ET.SubElement(data_write_accesss,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DWA_{pport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}'
    port_prototype_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/NvData/{If_name}/{DE}'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}


# data Receive Point By Argument or data Receive Point By Value and data send point >> Explicit

def drpa (): #completed
	global data_receive_point_by_arguments
	data_receive_point_by_arguments=ET.SubElement(runnable_entity,'DATA-RECEIVE-POINT-BY-ARGUMENTS')

def DRPA_RPort_SR_DataElement(currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed
    

    variable_access=ET.SubElement(data_receive_point_by_arguments,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DRP_{rport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
    port_prototype_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def DRPA_RPort_nvd_NvData(currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed 
 
    variable_access=ET.SubElement(data_receive_point_by_arguments,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DRP_{rport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
    port_prototype_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/NvData/{If_name}/{DE}'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}


def drpv (): #completed
	global data_receive_point_by_values
	data_receive_point_by_values=ET.SubElement(runnable_entity,'DATA-RECEIVE-POINT-BY-VALUES')

def DRPV_RPort_SR_DataElement(currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed
    

    variable_access=ET.SubElement(data_receive_point_by_values,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DRP_{rport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
    port_prototype_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def DRPV_RPort_nvd_NvData(currentfolder, CurrentSWC_shortname, rport, If_name, DE):#completed
 
	variable_access=ET.SubElement(data_receive_point_by_values,'VARIABLE-ACCESS')
	variable_access.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_access,'SHORT-NAME')
	short_name.text=f'DRP_{rport}_{DE}'
	accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
	autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
	port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
	port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
	port_prototype_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
	target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/NvData/{If_name}/{DE}'
	target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def dsp (): #completed   
    global data_send_points
    data_send_points=ET.SubElement(runnable_entity,'DATA-SEND-POINTS')

def DSP_PPort_SR_DataElement(currentfolder, CurrentSWC_shortname, pport, If_name, DE):#completed 

    variable_access=ET.SubElement(data_send_points,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DSP_{pport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/PPort_SR'
    port_prototype_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/SenderReceiver/{If_name}/{DE}' #'/SharedElements/PortInterfaces/SenderReceiver/SenderReceiverInterface/DataElement'
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}


def DSP_PPort_nvd_NvData(currentfolder, CurrentSWC_shortname, pport, If_name, DE):#completed 
 
    variable_access=ET.SubElement(data_send_points,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'DSP_{pport}_{DE}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    autosar_variable_iref=ET.SubElement(accessed_variable,'AUTOSAR-VARIABLE-IREF')
    port_prototype_ref=ET.SubElement(autosar_variable_iref,'PORT-PROTOTYPE-REF')
    port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}' 
    port_prototype_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
    target_data_prototype_ref=ET.SubElement(autosar_variable_iref,'TARGET-DATA-PROTOTYPE-REF')
    target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/NvData/{If_name}/{DE}' 
    target_data_prototype_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}





#-------------------------InterRunnableVariable------------------------------------#

def IRVRA (): #completed
	global read_local_variables
	read_local_variables=ET.SubElement(runnable_entity,'READ-LOCAL-VARIABLES')

def IRVRA_ExplicitInterRunnableVariable(ExplicitIRV_shortname, currentfolder, CurrentSWC_shortname ):#completed

    variable_access=ET.SubElement(read_local_variables,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'IRVRA_{ExplicitIRV_shortname}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    local_variable_ref=ET.SubElement(accessed_variable,'LOCAL-VARIABLE-REF')
    local_variable_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{ExplicitIRV_shortname}'
    local_variable_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def IRVRA_ImplicitInterRunnableVariable(ImplicitIRV_shortname, currentfolder, CurrentSWC_shortname):#completed 
 
	variable_access=ET.SubElement(read_local_variables,'VARIABLE-ACCESS')
	variable_access.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_access,'SHORT-NAME')
	short_name.text=f'IRVRA_{ImplicitIRV_shortname}'
	accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
	local_variable_ref=ET.SubElement(accessed_variable,'LOCAL-VARIABLE-REF')
	local_variable_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{ImplicitIRV_shortname}'
	local_variable_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}

def IRVWA (): #completed
    global written_local_variables
    written_local_variables=ET.SubElement(runnable_entity,'WRITTEN-LOCAL-VARIABLES')

def IRVWA_ExplicitInterRunnableVariable(ExplicitIRV_shortname, currentfolder, CurrentSWC_shortname):#completed

    variable_access=ET.SubElement(written_local_variables,'VARIABLE-ACCESS')
    variable_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(variable_access,'SHORT-NAME')
    short_name.text=f'IRVWA_{ExplicitIRV_shortname}'
    accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
    local_variable_ref=ET.SubElement(accessed_variable,'LOCAL-VARIABLE-REF')
    local_variable_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{ExplicitIRV_shortname}'
    local_variable_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}
 
def IRVWA_ImplicitInterRunnableVariable(ImplicitIRV_shortname, currentfolder, CurrentSWC_shortname):#completed
 
	variable_access=ET.SubElement(written_local_variables,'VARIABLE-ACCESS')
	variable_access.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(variable_access,'SHORT-NAME')
	short_name.text=f'IRVWA_{ImplicitIRV_shortname}'
	accessed_variable=ET.SubElement(variable_access,'ACCESSED-VARIABLE')
	local_variable_ref=ET.SubElement(accessed_variable,'LOCAL-VARIABLE-REF')
	local_variable_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{ImplicitIRV_shortname}'
	local_variable_ref.attrib={'DEST':'VARIABLE-DATA-PROTOTYPE'}



#-------------------------Mode------------------------------------#

def msp(): #completed
	global mode_switch_points

	mode_switch_points=ET.SubElement(runnable_entity,'MODE-SWITCH-POINTS')

def MSP_PPort_msi_ModeGroup(currentfolder, CurrentSWC_shortname, pport, If_name, modegroup):#completed


	mode_switch_point=ET.SubElement(mode_switch_points,'MODE-SWITCH-POINT')
	mode_switch_point.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(mode_switch_point,'SHORT-NAME')
	short_name.text=f'MSP_{pport}_{modegroup}'
	mode_group_iref=ET.SubElement(mode_switch_point,'MODE-GROUP-IREF')
	context_p_port_ref=ET.SubElement(mode_group_iref,'CONTEXT-P-PORT-REF')
	context_p_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{pport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/PPort_msi'
	context_p_port_ref.attrib={'DEST':'P-PORT-PROTOTYPE'}
	target_mode_group_ref=ET.SubElement(mode_group_iref,'TARGET-MODE-GROUP-REF')
	target_mode_group_ref.text=f'/SharedElements/PortInterfaces/ModeSwitch/{If_name}/{modegroup}' #'/SharedElements/PortInterfaces/ModeSwitch/ModeSwitchInterface/ModeGroup'
	target_mode_group_ref.attrib={'DEST':'MODE-DECLARATION-GROUP-PROTOTYPE'}

#-------------------------parameter------------------------------------#

def pa (): #completed
	global parameter_accesss
	parameter_accesss=ET.SubElement(runnable_entity,'PARAMETER-ACCESSS')

def CMCPA_ConstantMemory(currentfolder, CurrentSWC_shortname,ConstantMemory_shortname):#completed 
    

    parameter_access=ET.SubElement(parameter_accesss,'PARAMETER-ACCESS')
    parameter_access.attrib={'UUID':rng.generate_uuid()}
    short_name=ET.SubElement(parameter_access,'SHORT-NAME')
    short_name.text= f'CMCPA_{ConstantMemory_shortname}'
    accessed_parameter=ET.SubElement(parameter_access,'ACCESSED-PARAMETER')
    local_parameter_ref=ET.SubElement(accessed_parameter,'LOCAL-PARAMETER-REF')
    local_parameter_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{ConstantMemory_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/ConstantMemory'
    local_parameter_ref.attrib={'DEST':'PARAMETER-DATA-PROTOTYPE'}

def PICPVA_PerInstanceParameter(currentfolder, CurrentSWC_shortname,per_instance_parameters_shortname):#completed 
 
	parameter_access=ET.SubElement(parameter_accesss,'PARAMETER-ACCESS')
	parameter_access.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_access,'SHORT-NAME')
	short_name.text=f'PICPVA_{per_instance_parameters_shortname}'
	accessed_parameter=ET.SubElement(parameter_access,'ACCESSED-PARAMETER')
	local_parameter_ref=ET.SubElement(accessed_parameter,'LOCAL-PARAMETER-REF')
	local_parameter_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{per_instance_parameters_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/PerInstanceParameter'
	local_parameter_ref.attrib={'DEST':'PARAMETER-DATA-PROTOTYPE'}
 
def CPA_RPort_prm_Parameter(currentfolder, CurrentSWC_shortname, rport, If_name, Parameter_shortname):#completed
 
	parameter_access=ET.SubElement(parameter_accesss,'PARAMETER-ACCESS')
	parameter_access.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_access,'SHORT-NAME')
	short_name.text=f'CPA_{rport}_{Parameter_shortname}'
	accessed_parameter=ET.SubElement(parameter_access,'ACCESSED-PARAMETER')
	autosar_parameter_iref=ET.SubElement(accessed_parameter,'AUTOSAR-PARAMETER-IREF')
	port_prototype_ref=ET.SubElement(autosar_parameter_iref,'PORT-PROTOTYPE-REF')
	port_prototype_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/RPort_prm'
	port_prototype_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	target_data_prototype_ref=ET.SubElement(autosar_parameter_iref,'TARGET-DATA-PROTOTYPE-REF')
	target_data_prototype_ref.text=f'/SharedElements/PortInterfaces/Parameter/{If_name}/{Parameter_shortname}' #'/SharedElements/PortInterfaces/Parameter/ParameterInterface/Parameter'
	target_data_prototype_ref.attrib={'DEST':'PARAMETER-DATA-PROTOTYPE'}

def SCPVA_SharedParameter(currentfolder, CurrentSWC_shortname,SharedParameter_shortname):#completed 
 
	parameter_access=ET.SubElement(parameter_accesss,'PARAMETER-ACCESS')
	parameter_access.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_access,'SHORT-NAME')
	short_name.text=f'SCPVA_{SharedParameter_shortname}'
	accessed_parameter=ET.SubElement(parameter_access,'ACCESSED-PARAMETER')
	local_parameter_ref=ET.SubElement(accessed_parameter,'LOCAL-PARAMETER-REF')
	local_parameter_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{IB_shortname}/{SharedParameter_shortname}' #'/SwComponentTypes/ApplSWC/ApplicationSwComponentType/IB_Appl/SharedParameter'
	local_parameter_ref.attrib={'DEST':'PARAMETER-DATA-PROTOTYPE'}



#-------------------------cs access------------------------------------#
def sscp () : #completed
	global server_call_points

	server_call_points=ET.SubElement(runnable_entity,'SERVER-CALL-POINTS')


def SSCP_RPort_CS_Operation(currentfolder, CurrentSWC_shortname, rport, If_name, operation):#completed 
 

	synchronous_server_call_point=ET.SubElement(server_call_points,'SYNCHRONOUS-SERVER-CALL-POINT')
	synchronous_server_call_point.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(synchronous_server_call_point,'SHORT-NAME')
	short_name.text=f'SSCP_{rport}_{operation}'
	operation_iref=ET.SubElement(synchronous_server_call_point,'OPERATION-IREF')
	context_r_port_ref=ET.SubElement(operation_iref,'CONTEXT-R-PORT-REF')
	context_r_port_ref.text=f'/SwComponentTypes/{currentfolder}/{CurrentSWC_shortname}/{rport}'
	context_r_port_ref.attrib={'DEST':'R-PORT-PROTOTYPE'}
	target_required_operation_ref=ET.SubElement(operation_iref,'TARGET-REQUIRED-OPERATION-REF')
	target_required_operation_ref.text=f'/SharedElements/PortInterfaces/ClientServer/{If_name}/{operation}' #f'/SharedElements/PortInterfaces/ClientServer/ClientServerInterface/Operation1'
	target_required_operation_ref.attrib={'DEST':'CLIENT-SERVER-OPERATION'}
	timeout=ET.SubElement(synchronous_server_call_point,'TIMEOUT')
	timeout.text='0'



########## SW Component types ########### 

def ApplicationSwComponentType(ApplSWC_folder_elements,ApplSWC_shortname): #completed
	global application_sw_component_type
	application_sw_component_type=ET.SubElement(ApplSWC_folder_elements,'APPLICATION-SW-COMPONENT-TYPE')
	application_sw_component_type.attrib={'UUID':rng.generate_uuid()} #automatically rng to be generated and everytime need to check the uuid in the xml file
	ApplSWC_shortname1=ET.SubElement(application_sw_component_type,'SHORT-NAME')
	ApplSWC_shortname1.text= ApplSWC_shortname
 
def ComplexDeviceDriverSwComponentType(CddSWC_folder_elements, CddSWC_shortname ): #completed
	global complex_device_driver_sw_component_type
	complex_device_driver_sw_component_type=ET.SubElement(CddSWC_folder_elements,'COMPLEX-DEVICE-DRIVER-SW-COMPONENT-TYPE')
	complex_device_driver_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(complex_device_driver_sw_component_type,'SHORT-NAME')
	short_name.text=CddSWC_shortname
 
def CompositionSwComponentType(CompSWC_folder_elements,CompSWC_folder_short_name):  #completed
	global composition_sw_component_type
	composition_sw_component_type=ET.SubElement(CompSWC_folder_elements,'COMPOSITION-SW-COMPONENT-TYPE')
	composition_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(composition_sw_component_type,'SHORT-NAME')
	short_name.text=CompSWC_folder_short_name

def EcuAbstractionSwComponentType(EcuAbSWC_folder_elements,EcuAbSWC_folder_short_name): #completed
	global ecu_abstraction_sw_component_type
	ecu_abstraction_sw_component_type=ET.SubElement(EcuAbSWC_folder_elements,'ECU-ABSTRACTION-SW-COMPONENT-TYPE')
	ecu_abstraction_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(ecu_abstraction_sw_component_type,'SHORT-NAME')
	short_name.text=EcuAbSWC_folder_short_name

def NvBlockSwComponentType(NvDataSWC_folder_elements,NvDataSWC_folder_short_name): #completed
	global nv_block_sw_component_type
	nv_block_sw_component_type=ET.SubElement(NvDataSWC_folder_elements,'NV-BLOCK-SW-COMPONENT-TYPE')
	nv_block_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(nv_block_sw_component_type,'SHORT-NAME')
	short_name.text=NvDataSWC_folder_short_name

def ParameterSwComponentType(PrmSWC_folder_elements,PrmSWC_folder_short_name):  #completed
	global parameter_sw_component_type
	parameter_sw_component_type=ET.SubElement(PrmSWC_folder_elements,'PARAMETER-SW-COMPONENT-TYPE')
	parameter_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(parameter_sw_component_type,'SHORT-NAME')
	short_name.text=PrmSWC_folder_short_name

def SensorActuatorSwComponentType(SnsrActSWC_folder_elements,SnsrActSWC_folder_short_name): #completed
	global sensor_actuator_sw_component_type
	sensor_actuator_sw_component_type=ET.SubElement(SnsrActSWC_folder_elements,'SENSOR-ACTUATOR-SW-COMPONENT-TYPE')
	sensor_actuator_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(sensor_actuator_sw_component_type,'SHORT-NAME')
	short_name.text=SnsrActSWC_folder_short_name

def ServiceProxySwComponentType(SrvcPrxySWC_folder_elements,SrvcPrxySWC_folder_short_name ): #completed
	global service_proxy_sw_component_type
	service_proxy_sw_component_type=ET.SubElement(SrvcPrxySWC_folder_elements,'SERVICE-PROXY-SW-COMPONENT-TYPE')
	service_proxy_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(service_proxy_sw_component_type,'SHORT-NAME')
	short_name.text=SrvcPrxySWC_folder_short_name

def ServiceSwComponentType(SrvcSWC_folder_elements, SrvcSWC_folder_short_name): #completed
	global service_sw_component_type
	service_sw_component_type=ET.SubElement(SrvcSWC_folder_elements,'SERVICE-SW-COMPONENT-TYPE')
	service_sw_component_type.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(service_sw_component_type,'SHORT-NAME')
	short_name.text=SrvcSWC_folder_short_name


########## Systems ########### 

def Systems(Systems_folder_elements,Systems_folder_short_name):  #completed

	Systems_folder=ET.SubElement(Systems_folder_elements,'AR-PACKAGE')
	Systems_folder.attrib={'UUID':rng.generate_uuid()}
	short_name=ET.SubElement(Systems_folder,'SHORT-NAME')
	short_name.text=Systems_folder_short_name

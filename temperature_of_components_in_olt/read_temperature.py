import time
import logging
import datetime
import json
import xlsxwriter 
from snmplib.snmp import SnmpInterface
from snmplib.oltmibs import sinaSP5100FanSpeed, sinaBoardCpuTemperature, sinaBoardPonTemperature, sinaBoardPonChipTemperature

# sinaSP5100FanSpeed    =  ".1.3.6.1.4.1.54964.2.1.1.2.1.2.1.1"
# sinaBoardCpuTemperature    =  ".1.3.6.1.4.1.54964.4.1.1.1.1.3"
# sinaBoardPonTemperature    =  ".1.3.6.1.4.1.54964.4.1.1.1.1.4"
# sinaBoardPonChipTemperature    =  ".1.3.6.1.4.1.54964.4.1.1.1.1.5"
# sinaBoardSwitchChipTemperature    =  ".1.3.6.1.4.1.54964.4.1.1.1.1.6"


logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)


def join_oid(oid, *indexes):
    list_indexes = [i for i in indexes]
    list_indexes = [i for i in list_indexes[0]]
    for index in list_indexes:
        if index == 0:
            oid += '.0'
        elif index:
            oid = ".".join([oid, str(index)])
    return oid



def get_Cpu_temperature(snmp_interface, oid_name_Cpu_temp,*suffix_index):
    temperature_CPU = snmp_interface.snmp_get(join_oid(oid_name_Cpu_temp, suffix_index))
    logger.info('getting read CPU_TEMPERATURE COMPLETE')  
    return temperature_CPU['value']

def get_Pon_temperature(snmp_interface, oid_name_Pon_temp,*suffix_index):
    temperature_Pon = snmp_interface.snmp_get(join_oid(oid_name_Pon_temp, suffix_index))
    logger.info('getting read PON_TEMPERATURE COMPLETE')  
    return temperature_Pon['value']

def get_Pon_Chip_temperature(snmp_interface, oid_name_Pon_Chip_temp,*suffix_index):
    temperature_Pon_Chip = snmp_interface.snmp_get(join_oid(oid_name_Pon_Chip_temp, suffix_index))
    logger.info('getting read PON_CHIP_TEMPERATURE COMPLETE')  

    return temperature_Pon_Chip['value']    

def set_and_get_fan_speed(snmp_interface, oid_fan_speed, set_value, *suffix_index): 
    fan_speed_set_result = snmp_interface.snmp_set(join_oid(oid_fan_speed, suffix_index), set_value, "Integer")
    fan_speed = 0
    if "noError" == fan_speed_set_result["error"]:
        logger.info('Fan setting is succesful')    
        fan_sp = snmp_interface.snmp_get(join_oid(oid_fan_speed, suffix_index))
        if str(fan_sp['value'])==str(set_value):
            logger.info('Setting Fan speed after getting is succesful')  
            fan_speed = fan_sp['value']
        else:
            logger.info('Setting Fan speed after getting is Failed')  
            fan_speed = 0
    else:
        logger.info(f'Fan setting before getting has difficulty {fan_speed_set_result["error"]}')         
    return fan_speed


ip_address = input("Please Enter Shelf IP:")
shelfIndex = input("Please Enter card shelf number:")
slotIndex = input("Please Enter card slot number:")
delay = input("Please Enter time extent want to wait for measurment:")

snmp_interface = SnmpInterface(ip=ip_address, community="sina_private", version="2", port=161, timeout=20)

for speed_fan_set in [10,50,90]:
    for fan_index in range(1,5):
        fan_speed = set_and_get_fan_speed(snmp_interface, sinaSP5100FanSpeed, speed_fan_set, shelfIndex, fan_index)
    time.sleep(delay)
    cpu_tp = get_Cpu_temperature(snmp_interface, sinaBoardCpuTemperature, shelfIndex, slotIndex)
    pon_tp = get_Pon_temperature(snmp_interface, sinaBoardPonTemperature, shelfIndex, slotIndex)
    pon_chip_tp = get_Pon_Chip_temperature(snmp_interface, sinaBoardPonChipTemperature, shelfIndex, slotIndex)
    print(cpu_tp, pon_tp, pon_chip_tp)

    wb = xlsxwriter.Workbook("workbook.xlsx")
    worksheet = wb.add_worksheet()
    worksheet.write("A1", 'CPU Temperature')
    worksheet.write("B1", 'PON Temperature')
    worksheet.write("C1", 'PON Chip Temperature')
    worksheet.write("D1", 'SPEED FAN')
    worksheet.write(f"A{i+1}", f'{cpu_tp}')
    worksheet.write(f"B{i+1}", f'{pon_tp}')
    worksheet.write(f"C{i+1}", f'{pon_chip_tp}')
    worksheet.write(f"D{i+1}", f'{fan_speed}')

wb.close()    



import time
import logging
import datetime
import json
import xlsxwriter 
from snmplib.snmp import SnmpInterface
from snmplib.oltmibs import sinaSP5100FanSpeed, sinaBoardCpuTemperature, sinaBoardPonTemperature, sinaBoardPonChipTemperature
import openpyxl 

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
    return temperature_CPU['value']/1000

def get_Pon_temperature(snmp_interface, oid_name_Pon_temp,*suffix_index):
    temperature_Pon = snmp_interface.snmp_get(join_oid(oid_name_Pon_temp, suffix_index))
    logger.info('getting read PON_TEMPERATURE COMPLETE')  
    return temperature_Pon['value']/1000

def get_Pon_Chip_temperature(snmp_interface, oid_name_Pon_Chip_temp,*suffix_index):
    temperature_Pon_Chip = snmp_interface.snmp_get(join_oid(oid_name_Pon_Chip_temp, suffix_index))
    logger.info('getting read PON_CHIP_TEMPERATURE COMPLETE')  

    return temperature_Pon_Chip['value']/1000    

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
state = input("Please Enter number of your state:")
state = int(state)
snmp_interface = SnmpInterface(ip=ip_address, community="sina_private", version="2", port=161, timeout=20)


if state == 1:
    wb = xlsxwriter.Workbook("/home/zeinab/python_script/temperature_of_components_in_olt/workbook.xlsx")
    worksheet_fan_variation = wb.add_worksheet("sheet1")
    worksheet2 = wb.add_worksheet("sheet2")
    worksheet3 = wb.add_worksheet("sheet3")
    worksheet4 = wb.add_worksheet("sheet4")
    worksheet5 = wb.add_worksheet("sheet5")
    worksheet6 = wb.add_worksheet("sheet6")
    worksheet7 = wb.add_worksheet("sheet7")
    worksheet8 = wb.add_worksheet("sheet8")
    worksheet9 = wb.add_worksheet("sheet9")

    worksheet_fan_variation.write("A1", 'Fan Speed Variation With Fan, Door and 2 AC')
    worksheet_fan_variation.write("A2", 'CPU Temperature  C')
    worksheet_fan_variation.write("B2", 'PON Temperature  C')
    worksheet_fan_variation.write("C2", 'PON Chip Temperature  C')
    worksheet_fan_variation.write("D2", 'SPEED FAN (percent of 100)')

    for speed_fan_set in [10,40,90]:
        for fan_index in range(1,5):
            fan_speed = set_and_get_fan_speed(snmp_interface, sinaSP5100FanSpeed, speed_fan_set, shelfIndex, fan_index)    
        for i in range(2,50):    
            cpu_tp = get_Cpu_temperature(snmp_interface, sinaBoardCpuTemperature, shelfIndex, slotIndex)
            pon_tp = get_Pon_temperature(snmp_interface, sinaBoardPonTemperature, shelfIndex, slotIndex)
            pon_chip_tp = get_Pon_Chip_temperature(snmp_interface, sinaBoardPonChipTemperature, shelfIndex, slotIndex)
            print(cpu_tp, pon_tp, pon_chip_tp)
            if speed_fan_set == 10:
                worksheet_fan_variation.write(f"A{i+1}", f'{cpu_tp}')
                worksheet_fan_variation.write(f"B{i+1}", f'{pon_tp}')
                worksheet_fan_variation.write(f"C{i+1}", f'{pon_chip_tp}')
                worksheet_fan_variation.write(f"D{i+1}", f'{fan_speed}')
            elif speed_fan_set == 40:  
                worksheet_fan_variation.write(f"A{i+50}", f'{cpu_tp}')
                worksheet_fan_variation.write(f"B{i+50}", f'{pon_tp}')
                worksheet_fan_variation.write(f"C{i+50}", f'{pon_chip_tp}')
                worksheet_fan_variation.write(f"D{i+50}", f'{fan_speed}')  
            elif speed_fan_set == 90:  
                worksheet_fan_variation.write(f"A{i+100}", f'{cpu_tp}')
                worksheet_fan_variation.write(f"B{i+100}", f'{pon_tp}')
                worksheet_fan_variation.write(f"C{i+100}", f'{pon_chip_tp}')
                worksheet_fan_variation.write(f"D{i+100}", f'{fan_speed}')      
            time.sleep(int(delay))
    wb.close() 

if state != 1:
    workbook = openpyxl.load_workbook('/home/zeinab/python_script/temperature_of_components_in_olt/workbook.xlsx')
    for speed_fan_set in [10,40,90]:
        for fan_index in range(1,5):
            fan_speed = set_and_get_fan_speed(snmp_interface, sinaSP5100FanSpeed, speed_fan_set, shelfIndex, fan_index)               
        if state == 2:
            worksheet = workbook['sheet2']
            worksheet["A1"] = 'With Filter'
        if state == 3:
            worksheet = workbook['sheet3']
            worksheet["A1"] = 'No Filter With Door and 2 AC (one connected and another not connected), Some module'
        if state == 4:
            worksheet = workbook['sheet4']
            worksheet["A1"] = 'Without Door, with filter and 2 AC (one connected and another not connected), Some module'
        if state == 5:
            worksheet = workbook['sheet5']
            worksheet["A1"] = 'Without Door and filter, with 2 AC (one connected and another not connected), Some module' 

        if state == 6:
            worksheet = workbook['sheet6']
            worksheet["A1"] = 'Full Module (4up 1pon) with Filter and Door, with 2 AC (one connected and another not connected)'
        if state == 7:
            worksheet = workbook['sheet7']
            worksheet["A1"] = 'Full Module (8up 1pon) without Filter With Door, with 2 AC (one connected and another not connected)'

        if state == 8:
            worksheet = workbook['sheet8']
            worksheet["A1"] = 'Full Module (1up 4pon) with Filter and Door, with 2 AC (one connected and another not connected)' 
        if state == 9:
            worksheet = workbook['sheet9']
            worksheet["A1"] = 'Full Module (1up 8pon)  with Filter and Door, with 2 AC (one connected and another not connected)' 
        if state == 10:
            worksheet = workbook['sheet10']
            worksheet["A1"] = 'Full Module (1up 12pon)  with Filter and Door, with 2 AC (one connected and another not connected)' 
        if state == 11:
            worksheet = workbook['sheet11']
            worksheet["A1"] = 'Full Module (1up 16pon)  with Filter and Door, with 2 AC (one connected and another not connected)' 

        if state == 12:
            worksheet = workbook['sheet12']
            worksheet["A1"] = 'Full Module (8up 16pon)  with Filter and Door, with 2 AC (one connected and another not connected)' 



        if state == 13:
            worksheet = workbook['sheet13']
            worksheet["A1"] = 'TwO AC, Some module with Door With Filter '
        if state == 14:
            worksheet = workbook['sheet14']
            worksheet["A1"] = 'Two AC, Some Module with Door Without Filter'
        if state == 15:
            worksheet = workbook['sheet15']
            worksheet["A1"] = 'One AC, Some Module with Door With Filter '
        if state == 16:
            worksheet = workbook['sheet16']
            worksheet["A1"] = 'One AC, Some Module with Door Without Filter'



        worksheet["A2"] = 'CPU Temperature  C'
        worksheet["B2"] = 'PON Temperature  C'
        worksheet["C2"] = 'PON Chip Temperature  C'
        worksheet["D2"] = 'SPEED FAN (percent of 100)'  
        for i in range(2,50):    
            cpu_tp = get_Cpu_temperature(snmp_interface, sinaBoardCpuTemperature, shelfIndex, slotIndex)
            pon_tp = get_Pon_temperature(snmp_interface, sinaBoardPonTemperature, shelfIndex, slotIndex)
            pon_chip_tp = get_Pon_Chip_temperature(snmp_interface, sinaBoardPonChipTemperature, shelfIndex, slotIndex)
            print(cpu_tp, pon_tp, pon_chip_tp)
            if speed_fan_set == 10:
                worksheet[f"A{i+1}"] = f'{cpu_tp}'
                worksheet[f"B{i+1}"] = f'{pon_tp}'
                worksheet[f"C{i+1}"] = f'{pon_chip_tp}'
                worksheet[f"D{i+1}"] = f'{fan_speed}'
            if speed_fan_set == 40:
                worksheet[f"A{i+50}"] = f'{cpu_tp}'
                worksheet[f"B{i+50}"] = f'{pon_tp}'
                worksheet[f"C{i+50}"] = f'{pon_chip_tp}'
                worksheet[f"D{i+50}"] = f'{fan_speed}'
            if speed_fan_set == 90:
                worksheet[f"A{i+100}"] = f'{cpu_tp}'
                worksheet[f"B{i+100}"] = f'{pon_tp}'
                worksheet[f"C{i+100}"] = f'{pon_chip_tp}'
                worksheet[f"D{i+100}"] = f'{fan_speed}'    
            time.sleep(int(delay))

    workbook.save('/home/zeinab/python_script/temperature_of_components_in_olt/workbook.xlsx')
       





      
    
     







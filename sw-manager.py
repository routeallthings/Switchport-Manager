#!/usr/bin/env python

'''
---AUTHOR---
Name: Matt Cross
Email: routeallthings@gmail.com

---PREREQ---
INSTALL netmiko (pip install netmiko)
INSTALL textfsm (pip install textfsm)
INSTALL openpyxl (pip install openpyxl)
INSTALL fileinput (pip install fileinput)
INSTALL xlhelper (python -m pip install git+git://github.com/routeallthings/xlhelper.git)
'''

#Module Imports (Native)
import re
import getpass
import os
import unicodedata
import csv
import threading
import time
import sys
import requests
from datetime import datetime

#Module Imports (Non-Native)
try:
	import netmiko
	from netmiko import ConnectHandler
except ImportError:
	netmikoinstallstatus = fullpath = raw_input ('Netmiko module is missing, would you like to automatically install? (Y/N): ')
	if "Y" in netmikoinstallstatus.upper() or "YES" in netmikoinstallstatus.upper():
		os.system('python -m pip install netmiko')
		import netmiko
		from netmiko import ConnectHandler
	else:
		print "You selected an option other than yes. Please be aware that this script requires the use of netmiko. Please install manually and retry"
		sys.exit()
try:
	from openpyxl import load_workbook
	from openpyxl import workbook
	from openpyxl import Workbook
	from openpyxl.styles import Font, NamedStyle
except ImportError:
	openpyxlinstallstatus = raw_input ('openpyxl module is missing, would you like to automatically install? (Y/N): ')
	if 'y' in openpyxlinstallstatus.lower():
		os.system('python -m pip install openpyxl')
		from openpyxl import load_workbook
		from openpyxl import workbook
		from openpyxl import Workbook
		from openpyxl.styles import Font, NamedStyle
	else:
		print 'You selected an option other than yes. Please be aware that this script requires the use of openpyxl. Please install manually and retry'
		print 'Exiting in 5 seconds'
		time.sleep(5)
		sys.exit()		
#
try:
	import fileinput
except ImportError:
	requestsinstallstatus = fullpath = raw_input ('FileInput module is missing, would you like to automatically install? (Y/N): ')
	if 'Y' in requestsinstallstatus or 'y' in requestsinstallstatus or 'yes' in requestsinstallstatus or 'Yes' in requestsinstallstatus or 'YES' in requestsinstallstatus:
		os.system('python -m pip install FileInput')
		import FileInput
	else:
		print 'You selected an option other than yes. Please be aware that this script requires the use of FileInput. Please install manually and retry'
		sys.exit()
# Darth-Veitcher Module https://github.com/darth-veitcher/xlhelper		
from pprint import pprint
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, column_index_from_string
from collections import OrderedDict
try:
	import xlhelper
except ImportError:
	requestsinstallstatus = fullpath = raw_input ('xlhelper module is missing, would you like to automatically install? (Y/N): ')
	if 'Y' in requestsinstallstatus or 'y' in requestsinstallstatus or 'yes' in requestsinstallstatus or 'Yes' in requestsinstallstatus or 'YES' in requestsinstallstatus:
		os.system('python -m pip install git+git://github.com/routeallthings/xlhelper.git')
		import xlhelper
	else:
		print 'You selected an option other than yes. Please be aware that this script requires the use of xlhelper. Please install manually and retry'
		sys.exit()
#######################################
#Functions
def DOWNLOAD_FILE(url,saveas):
    # NOTE the stream=True parameter
    r = requests.get(url, stream=True)
    with open(saveas, 'wb') as f:
        for chunk in r.iter_content(chunk_size=1024): 
            if chunk: # filter out keep-alive new chunks
                f.write(chunk)
def RestartPort(device,vendormac):
	deviceip = device.get('IP').encode('utf-8')
	devicevendor = device.get('Vendor').encode('utf-8')
	devicetype = device.get('Type').encode('utf-8')
	devicetype = devicevendor.lower() + "_" + devicetype.lower()
	#Start Connection
	try:
		sshnet_connect = ConnectHandler(device_type=devicetype, ip=deviceip, username=sshusername, password=sshpassword, secret=enablesecret)
		devicehostname = sshnet_connect.find_prompt()
		devicehostname = devicehostname.strip('#')
		if '>' in devicehostname:
			sshnet_connect.enable()
			devicehostname = devicehostname.strip('>')
			devicehostname = sshnet_connect.find_prompt()
			devicehostname = devicehostname.strip('#')
		vendormac = vendormac[:4] + '.' + vendormac[4:]
		showvendormac = 'show mac address-table | include ' + vendormac
		fullmaclist = sshnet_connect.send_command(showvendormac)
		if fullmaclist == '':
			sys.exit()
		fullmaclist = fullmaclist.split('\n')
		# Get Port #
		portlist = []
		for mac in fullmaclist:
			mac = " ".join(mac.split())
			maclist = mac.split(' ')
			portlist.append(maclist[3])
		# Restart Port #
		sshcommandset = []
		for port in portlist:
			iport = 'interface ' + port
			ishut = 'shutdown'
			inoshut = 'no shutdown'
			sshcommandset.append(iport)
			sshcommandset.append(ishut)
			sshcommandset.append(inoshut)
		FullOutput = sshnet_connect.send_config_set(sshcommandset)
		OutputLogp = loglocation + '\\' + devicehostname + '_portrestart.txt'
		if os.path.exists(OutputLogp):
			OutputLog = open(OutputLogp,'a+')
		else:
			OutputLog = open(OutputLogp,'w+')
		OutputLog.write('#################################################################\n')
		OutputLog.write('Start of Log\n')
		OutputLog.write('Current Start Time: ' + str(datetime.now()) + '\n')
		OutputLog.write(FullOutput)
		OutputLog.write('\n')
		OutputLog.close()
		sshnet_connect.disconnect()
	except Exception as e:
		print 'Error while gathering data with ' + deviceip + '. Error is ' + str(e)
		try:
			sshnet_connect.disconnect()
		except:
			'''Nothing'''
	except KeyboardInterrupt:
		print 'CTRL-C pressed, exiting update of DB'
		try:
			sshnet_connect.disconnect()
		except:
			'''Nothing'''

def SetVLAN(device,vendormac,vendorvlan,vendorvlanmod):
	deviceip = device.get('IP').encode('utf-8')
	devicevendor = device.get('Vendor').encode('utf-8')
	devicetype = device.get('Type').encode('utf-8')
	devicetype = devicevendor.lower() + "_" + devicetype.lower()
	#Start Connection
	try:
		sshnet_connect = ConnectHandler(device_type=devicetype, ip=deviceip, username=sshusername, password=sshpassword, secret=enablesecret)
		devicehostname = sshnet_connect.find_prompt()
		devicehostname = devicehostname.strip('#')
		if '>' in devicehostname:
			sshnet_connect.enable()
			devicehostname = devicehostname.strip('>')
			devicehostname = sshnet_connect.find_prompt()
			devicehostname = devicehostname.strip('#')
		vendormac = vendormac[:4] + '.' + vendormac[4:]
		showvendormac = 'show mac address-table | include ' + vendormac
		fullmaclist = sshnet_connect.send_command(showvendormac)
		if fullmaclist == '':
			sys.exit()
		fullmaclist = fullmaclist.split('\n')
		# Get Port #
		vlanportlist = []
		for mac in fullmaclist:
			mac = " ".join(mac.split())
			maclist = mac.split(' ')
			vlanport = maclist[0] + ',' + maclist[3]
			vlanportlist.append(vlanport)
		# Build list of changes
		findtrunk = 'show interface trunk | i trunking'
		trunklist = sshnet_connect.send_command(findtrunk)
		sshcommandset = []
		portcompliance = []
		for vlanport in vlanportlist:
			vlanport = vlanport.split(',')
			vlannumber = vlanport[0]
			vlaninterface = vlanport[1]
			if vlaninterface in trunklist:
				notrunk = 0
			else:
				notrunk = 1
			if not vlannumber in vendorvlan:
				if modifyvlan == 0 and notrunk == 1:
					print 'Port is out of compliance: ' + devicehostname + ',' + vlaninterface + ',' + vlannumber
				if modifyvlan == 1 and notrunk == 1:
					iport = 'interface ' + vlaninterface
					ivlan = 'switchport access vlan ' + vendorvlanmod
					sshcommandset.append(iport)
					sshcommandset.append(ivlan)
		if modifyvlan == 1:
			FullOutput = sshnet_connect.send_config_set(sshcommandset)
			OutputLogp = loglocation + '\\' + devicehostname + '_portvlan.txt'
			if os.path.exists(OutputLogp):
				OutputLog = open(OutputLogp,'a+')
			else:
				OutputLog = open(OutputLogp,'w+')
			OutputLog.write('#################################################################\n')
			OutputLog.write('Start of Log\n')
			OutputLog.write('Current Start Time: ' + str(datetime.now()) + '\n')
			OutputLog.write(FullOutput)
			OutputLog.write('\n')
			OutputLog.close()
		sshnet_connect.disconnect()
	except Exception as e:
		print 'Error with sending commands to ' + deviceip + '. Error is ' + str(e)
		try:
			sshnet_connect.disconnect()
		except:
			'''Nothing'''
	except KeyboardInterrupt:
		print 'CTRL-C pressed, exiting update of switches'
		try:
			sshnet_connect.disconnect()
		except:
			'''Nothing'''

def ExportVLANs(device):
	deviceip = device.get('IP').encode('utf-8')
	devicevendor = device.get('Vendor').encode('utf-8')
	devicetype = device.get('Type').encode('utf-8')
	devicetype = devicevendor.lower() + "_" + devicetype.lower()
	#Start Connection
	try:
		sshnet_connect = ConnectHandler(device_type=devicetype, ip=deviceip, username=sshusername, password=sshpassword, secret=enablesecret)
		devicehostname = sshnet_connect.find_prompt()
		devicehostname = devicehostname.strip('#')
		if '>' in devicehostname:
			sshnet_connect.enable()
			devicehostname = devicehostname.strip('>')
			devicehostname = sshnet_connect.find_prompt()
			devicehostname = devicehostname.strip('#')
		devicehostnames.append(devicehostname)
		showinterfacelist = 'sh ip int br | e Vlan|Tunnel|Loopback|Port-channel|GigabitEthernet0/0|Interface'
		fullinterfacelist = sshnet_connect.send_command(showinterfacelist)
		if fullinterfacelist == '':
			sys.exit()
		fullinterfacelist = fullinterfacelist.split('\n')
		for fullint in fullinterfacelist:
			interfacedictionary = {}
			fullintreg = re.search('(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)',fullint)
			try:
				intname = fullintreg.group(1)
			except:
				intname = ''
			try:
				intstatus = fullintreg.group(5)
			except:
				intstatus = ''
			if 'up' in intstatus:
				showmaccmd = 'show mac address-table interface ' + intname + ' | include /'
				showmac = sshnet_connect.send_command(showmaccmd)
				showmac = re.search('(\S+)\s+(\S+)\s+(\S+)\s+(\S+)',showmac)
				try:
					intmac = showmac.group(2)
				except:
					intmac = 'No MAC or Error'
			else:
				intmac = ''
			showrunintcmd = 'show running-config interface ' + intname
			try:
				showrunint = sshnet_connect.send_command(showrunintcmd)
				showrunint = showrunint.split('\n')
				intvlan = ''
				inttemplate = ''
				for int in showrunint:
					if 'switchport access' in int:
						intvlan = re.search('(\S+)\s+(\S+)\s+(\S+)\s+(\S+)',int)
						try:
							intvlan = intvlan.group(4)
						except:
							intvlan = 'Error'
					if 'switchport mode trunk' in int:
						intvlan = 'Trunk'
					if 'source template' in int:
						inttemplate = re.search('(\S+)\s+(\S+)\s+(\S+)',int)
						try:
							inttemplate = inttemplate.group(3)
						except:
							inttemplate = ''
			except:
				intvlan = ''
				inttemplate = ''
			interfacedictionary['Hostname'] = devicehostname
			interfacedictionary['Interface'] = intname
			interfacedictionary['VLAN'] = intvlan
			interfacedictionary['Status'] = intstatus
			interfacedictionary['MacAddress'] = intmac
			interfacedictionary['Template'] = inttemplate
			finalinterfacelist.append(interfacedictionary) 
		sshnet_connect.disconnect()
	except Exception as e:
		print 'Error with sending commands to ' + deviceip + '. Error is ' + str(e)
		try:
			sshnet_connect.disconnect()
		except:
			'''Nothing'''
	except KeyboardInterrupt:
		print 'CTRL-C pressed, exiting update of switches'
		try:
			sshnet_connect.disconnect()
		except:
			'''Nothing'''

#########################################
print ''
print 'Switchport Manager'
print '############################################################'
print 'The purpose of this tool is to use a XLSX import to update'
print 'a ports vlan based on the vendor mac address and/or restart'
print 'the port based on the mac address.'
print 'Please fill in the config tab on the templated XLSX sheet.'
print '############################################################'
print ''
print '----Questions that need answering----'
excelfilelocation = raw_input('File to load the excel data from (e.g. C:/Python27/sw-config.xlsx):')
if excelfilelocation == '':
	excelfilelocation = 'C:/Python27/sw-config.xlsx'
excelfilelocation = excelfilelocation.replace('"', '')
# Load Configuration Variables
configdict = {}
for configvariables in xlhelper.sheet_to_dict(excelfilelocation,'Config'):
	try:
		configvar = configvariables.get('Variable').encode('utf-8')
		configval = configvariables.get('Value').encode('utf-8')
	except:
		configvar = configvariables.get('Variable')
		configval = configvariables.get('Value')
	configdict[configvar] = configval
# Username Variables/Questions
sshusername = configdict.get('Username')
if 'NA' == sshusername:
	sshusername = raw_input('What is the username you will use to login to the devices?:')
sshpassword = configdict.get('Password')
if 'NA' == sshpassword:
	sshpassword = getpass.getpass('What is the password you will use to login to the devices?:')
enablesecret = configdict.get('EnableSecret')
if 'NA' == enablesecret:
	enablesecret = getpass.getpass('What is the enable password you will use to access the devices?:')
# Rest of the Config Variables
loglocation = configdict.get('LogLocation')
if loglocation == None:
	loglocation = r'C:\Scripts\SWManager\Log'
vendormac = configdict.get('VendorMAC')
if vendormac == None:
	vendormac = raw_input('Please enter the first 6 characters of the vendor mac you want to match on?')
vendorvlan = configdict.get('VendorVLAN')
if vendorvlan == None:
	vendorvlan = raw_input('Please enter the vlan number you want to modify the ports to?')
vendorvlan = str(vendorvlan)
if ',' in vendorvlan:
	vendorvlan = vendorvlan.split(',')
vendorvlanmod = vendorvlan
exportlocation = configdict.get('ExportLocation')
if exportlocation == None:
	exportlocation = raw_input('Please enter the file path for the report (e.g. C:\Scripts\SWManager):')
	if exportlocation == '':
		exportlocation = 'C:\Scripts\SWManager'
exportlocation = str(exportlocation)
# Report Header Style
HeaderFont = Font(bold=True)
HeaderFont.size = 12
HeaderStyle = NamedStyle(name='BoldHeader')
HeaderStyle.font = HeaderFont
#
maclookupdburl = "http://standards-oui.ieee.org/oui.txt"
modifyvlan = 0
devicehostnames = []
finalinterfacelist = []
devicelist = []
for devices in xlhelper.sheet_to_dict(excelfilelocation,'Device IPs'):
	devicelist.append(devices)
# Create Log folder if its missing
newinstall = 0
if not os.path.exists(loglocation):
	os.makedirs(loglocation)
	newinstall = 1
#### Menu 
print ''
print '#################################################'
print '###                                           ###'
print '###        Please select an option below      ###'
print '###                                           ###'
print '###  1. Restart all vendor mac ports          ###'
print '###  2. Report VLAN for all vendor mac ports  ###'
print '###  3. Update VLAN for all vendor mac ports  ###'
print '###  4. Change the vendor mac address         ###'
print '###  5. Export Report to CSV of assignment    ###'
print '###  6. Exit                                  ###'
print '###                                           ###'
print '#################################################'
print ''
menuoption = raw_input('Selection (1-6)?:')
if menuoption == '':
	menuoption = 6
menuoption = int(menuoption)
# Modify VLAN option
if menuoption == 3:
	modifyvlan = 1
	if isinstance(vendorvlan, list):
		print '**************************************************'
		print '* Please select a VLAN in the list to change to: *'
		print '**************************************************'
		rowc = 1
		for v in vendorvlan:
			print str(rowc) + '. ' + v
			rowc = rowc + 1
		vendorvlanmod = raw_input('Please enter the vlan number you want to modify the ports to?: ')
else:
	modifyvlan = 0
if menuoption == 4:
		vendormac = raw_input('Please enter the first 6 characters of the vendor mac you want to match on?: ')
# Start
while (menuoption < 6):
	if menuoption == 1:
		if __name__ == "__main__":
			# Start Threads
			print 'Starting restart of all devices based on vendor mac'
			for device in devicelist:	
				deviceip = device.get('IP').encode('utf-8')
				t = threading.Thread(target=RestartPort, args=(device,vendormac))
				t.start()
			main_thread = threading.currentThread()
			# Join All Threads
			for it_thread in threading.enumerate():
				if it_thread != main_thread:
					it_thread.join()
		print 'Successfully restarted all ports based on vendor mac'
	if menuoption == 2 or menuoption == 3:
		if __name__ == "__main__":
		# Start Threads
			print 'Starting comparison/update of switchport based on vendor mac'
			for device in devicelist:	
				deviceip = device.get('IP').encode('utf-8')
				t = threading.Thread(target=SetVLAN, args=(device,vendormac,vendorvlan,vendorvlanmod))
				t.start()
			main_thread = threading.currentThread()
			# Join All Threads
			for it_thread in threading.enumerate():
				if it_thread != main_thread:
					it_thread.join()
		print 'Successfully changed or viewed the vlan on all ports based on vendor mac'
	if menuoption == 5:
		if __name__ == "__main__":
		# Start Threads
			print 'Starting export to CSV of all switchports'
			for device in devicelist:	
				deviceip = device.get('IP').encode('utf-8')
				t = threading.Thread(target=ExportVLANs, args=(device,))
				t.start()
			main_thread = threading.currentThread()
			# Join All Threads
			for it_thread in threading.enumerate():
				if it_thread != main_thread:
					it_thread.join()
			try:
				maclookupfilename = 'oui.txt'
				if not os.path.isfile(maclookupfilename):
					DOWNLOAD_FILE(maclookupdburl, maclookupfilename)
				maclookupdbo = open(maclookupfilename)
				maclookupdb = maclookupdbo.readlines()
				maclookupdbo.close()
				skipmac = 0
			except Exception as e:
				skipmac = 1
				print 'Could not load MAC database. Error is ' + str(e)		
			try:
				wb = Workbook()
				dest_filename = 'VLAN-Report.xlsx'
				dest_path = exportlocation + '\\' + dest_filename
				# Multiple Devices Report (Separate Tabbed) WIP
				dcount = 0
				for d in devicehostnames:
					dcount = dcount + 1
				#
				ws1 = wb.active
				# Continue on with work
				ws1.title = "VLAN Export"
				ws1.append(['Hostname','Interface','VLAN','Port Status','MacAddress','MacVendor','Template'])
				startrow = 2
				for row in finalinterfacelist:
					vendormac = row.get('MacAddress')
					if not vendormac == '':
						try:
							mac_company_mac = str(vendormac[0:7].replace('.','')).upper()
							for line in maclookupdb:
								if line.startswith(mac_company_mac):
									linev = line.replace('\n','').replace('\t',' ')
									maccompany = re.search(r'^[A-Z0-9]{6}\s+\(base 16\)\s+(.*)',linev).group(1)
								if maccompany == '' or maccompany == None:
									maccompany = 'Unknown'
						except:
							maccompany = 'Unknown'
					else:
						maccompany = ''
					# Add to workbook
					ws1['A' + str(startrow)] = row.get('Hostname')
					ws1['B' + str(startrow)] = row.get('Interface')
					ws1['C' + str(startrow)] = row.get('VLAN')
					ws1['D' + str(startrow)] = row.get('Status')
					ws1['E' + str(startrow)] = row.get('MacAddress')
					ws1['F' + str(startrow)] = maccompany
					ws1['G' + str(startrow)] = row.get('Template')
					startrow = startrow + 1
				wb.add_named_style(HeaderStyle)
				# Set styles on header row
				for cell in ws1["1:1"]:
					cell.style = 'BoldHeader'
				wb.save(filename = dest_path)
				wb.save(filename = dest_path)
				print 'Successfully created VLAN Report'
			except Exception as e:
				print 'Error creating VLAN Report. Error is ' + str(e)
		print 'Successfully exported vlans on all ports to xlsx'
	print ''
	print '#################################################'
	print '###                                           ###'
	print '###        Please select an option below      ###'
	print '###                                           ###'
	print '###  1. Restart all vendor mac ports          ###'
	print '###  2. Report VLAN for all vendor mac ports  ###'
	print '###  3. Update VLAN for all vendor mac ports  ###'
	print '###  4. Change the vendor mac address         ###'
	print '###  5. Export Report to CSV of assignment    ###'
	print '###  6. Exit                                  ###'
	print '###                                           ###'
	print '#################################################'
	print ''
	menuoption = raw_input('Selection (1-5)?:')
	if menuoption == '':
		menuoption = 5
	menuoption = int(menuoption)
	if menuoption == 3:
		modifyvlan = 1
		if isinstance(vendorvlan, list):
			print '**************************************************'
			print '* Please select a VLAN in the list to change to: *'
			print '**************************************************'
			rowc = 1
			for v in vendorvlan:
				print str(rowc) + '. ' + v
				rowc + 1
			vendorvlanmod = raw_input('Please enter the vlan number you want to modify the ports to?: ')
	else:
		modifyvlan = 0
	if menuoption == 4:
		vendormac = raw_input('Please enter the first 6 characters of the vendor mac you want to match on?: ')
print ''
print 'Exiting...'
print 'Thanks for playing'
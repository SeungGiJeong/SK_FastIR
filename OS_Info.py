import wmi
import Def_Locale
import logging
import logging.handlers

def Calc_UTC_Min(min):
	temp = min%60
	if temp==0:
		return "00"
	else:
		return str(temp)

def Parse_Datetime(datetime):
	return datetime[0:4]+"/"+datetime[4:6]+"/"+datetime[6:8]+" "+datetime[8:10]+":"+datetime[10:12]+":"+datetime[12:]


def extract_datetimeandUTC(datetimeandUTC):
	time = Parse_Datetime(datetimeandUTC.split('+')[0])
	UTC_H = int(datetimeandUTC.split('+')[1])/60
	UTC_M = Calc_UTC_Min(int(datetimeandUTC.split('+')[1]))
	return time+"(UTC+"+str(UTC_H)+":"+str(UTC_M)+")"

def get_process_info(para_options):
	wmi_pointer = wmi.WMI()
	logger=para_options["logger"]
	codepage=wmi_pointer.Win32_OperatingSystem(["CodeSet"])[0].CodeSet
	logger.info("#System INFO")
	logger.info(str("  *Windows Version\t:"+wmi_pointer.Win32_OperatingSystem(["Caption"])[0].Caption))
	logger.info(str("  *Windows BuildNumber:" + wmi_pointer.Win32_OperatingSystem(["BuildNumber"])[0].BuildNumber))

	logger.info(str("  *Windows Setup time\t:"+ extract_datetimeandUTC(wmi_pointer.Win32_OperatingSystem(["InstallDate"])[0].InstallDate)))
	logger.info("#System")
	logger.info(str("  *Processor Name\t:"+ wmi_pointer.Win32_Processor(["Name"])[0].Name))
	logger.info(str("  *Processor ID \t:"+ wmi_pointer.Win32_Processor(["ProcessorID"])[0].ProcessorID))
#	logger.info(str("  *Processor Load Info\t:"+ wmi_pointer.Win32_Processor(["LoadPercentage"])[0].LoadPercentage))
	logger.info(str("  *Physical_Mem Size\t:"+ str(int(wmi_pointer.Win32_PhysicalMemory(["Capacity"])[0].Capacity)/1048576/1024)+ "GByte"))
#	logger.info(str("  *Vitual_Mem Size\t:"+ int(wmi_pointer.Win32_OperatingSystem(["TotalVirtualMemorySize"])[0].TotalVirtualMemorySize)/1048576+ "Byte"))
	logger.info(str("  *System Type(86/64)\t:"+ wmi_pointer.Win32_OperatingSystem(["OSArchitecture"])[0].OSArchitecture.encode('ascii', 'ignore')))
	logger.info(str("  *Computer name\t:"+ wmi_pointer.Win32_OperatingSystem(["CSName"])[0].CSName))
	logger.info(str("#User"))
	logger.info(str("  *User count\t\t:"+str(wmi_pointer.Win32_OperatingSystem(["NumberOfUsers"])[0].NumberOfUsers)))
	logger.info(str("  *User name\t\t:"+wmi_pointer.Win32_OperatingSystem(["RegisteredUser"])[0].RegisteredUser))
	logger.info(str("#Detail Info"))
	logger.info(str("  *CodePage\t\t:"+ codepage))
	logger.info(str("  *CountryCode\t:"+ wmi_pointer.Win32_OperatingSystem(["CountryCode"])[0].CountryCode))
	logger.info(str("  *Locale\t\t:"+ Def_Locale.List[hex(int(wmi_pointer.Win32_OperatingSystem(["Locale"])[0].Locale,16))[2:]]))
	logger.info(str("  *MUILanguage\t:"+ wmi_pointer.Win32_OperatingSystem(["MUILanguages"])[0].MUILanguages[0]))
	logger.info(str("  *UTC\t\t:+"+ str(int(wmi_pointer.Win32_OperatingSystem(["CurrentTimeZone"])[0].CurrentTimeZone)/60)+ ":"+ Calc_UTC_Min(int(wmi_pointer.Win32_OperatingSystem(["CurrentTimeZone"])[0].CurrentTimeZone))))
	logger.info(str("  *Serial Number\t:"+wmi_pointer.Win32_OperatingSystem(["SerialNumber"])[0].SerialNumber ))
	logger.info(str("  *Last BootUP time\t:"+extract_datetimeandUTC(wmi_pointer.Win32_OperatingSystem(["LastBootUpTime"])[0].LastBootUpTime)))
	logger.info(str("  *System Drive\t:"+wmi_pointer.Win32_OperatingSystem(["SystemDrive"])[0].SystemDrive))
	logger.info(str("  *System Dircetory\t:"+wmi_pointer.Win32_OperatingSystem(["SystemDirectory"])[0].SystemDirectory))
	logger.info(str("  *Boot Device\t:"+wmi_pointer.Win32_OperatingSystem(["BootDevice"])[0].BootDevice))
	logger.info(str("  *Process Cnt\t:"+str(wmi_pointer.Win32_OperatingSystem(["NumberOfProcesses"])[0].NumberOfProcesses)))
	logger.info(str("#Storage"))
	logger.info(str("  *Number of Storgae\t:"+ str(len(wmi_pointer.Win32_DiskDrive(["Caption"])))))
	logger.info(str("  *Number of Partition:"+str(len(wmi_pointer.Win32_DiskDriveToDiskPartition(["Dependent"])))))
	logger.info(str("  *Number of Drive\t:"+ str(len(wmi_pointer.Win32_LogicalDiskToPartition(["Dependent"])))))
	logger.info(str("#Storage_detail"))
	logger.info(str("  *Storage"))
	for i in range(0, len(wmi_pointer.Win32_DiskDrive(["Caption"]))):
		logger.info(str("     Disk["+str(i)+"] "+ str(wmi_pointer.Win32_DiskDrive(["Caption"])[i].Caption)))
		logger.info(str("        Size\t\t:"+str(int((wmi_pointer.Win32_DiskDrive(["Size"])[i].Size))/1024/1024/1024)+ "GByte"))
		logger.info(str("        Serial Number\t:"+str(wmi_pointer.Win32_DiskDrive(["SerialNumber"])[i].SerialNumber.strip())))
		for j in range(0, wmi_pointer.Win32_DiskDrive(["Partitions"])[i].Partitions):
			drive_name=""
			try:
				drive_name=wmi_pointer.Win32_DiskDrive()[i].associators("Win32_DiskDriveToDiskPartition")[j].associators("Win32_LogicalDiskToPartition")[0].Caption
			except:
				pass
			logger.info(str("              Disk["+str(i)+"] Partition["+str(j)+"]+ Drive["+drive_name+"]"))
	logger.info(str("  *Logical Drive"	))	
	for i in range(0, len(wmi_pointer.Win32_LogicalDisk())):
		logger.info(str("     #Drive["+ wmi_pointer.Win32_LogicalDisk()[i].DeviceID+"]    size:"+ str(int(wmi_pointer.Win32_LogicalDisk()[i].size)/1024/1024/1024)+ "GByte"))


' VB Script to display computer details
' Created by Jack Henry
' https://github.com/JackInSights/

On Error Resume Next

' Define Local Variables
' ######################

Dim getWMI_obj
Dim getCompSettings
Dim getOSSettings
Dim getBIOSSettings
Dim getComputer
Dim getWMI_str
Dim HeadlineInfo
Dim ipNum
Dim monNum

Const vbLongDate = 1


' # Define queries
' #################
'Set getWMI_obj = GetObject("winmgmts:")

Set getWMI_obj = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Set getCompSettings = getWMI_obj.ExecQuery ("Select * from Win32_ComputerSystem")
Set getOSSettings   = getWMI_obj.ExecQuery ("Select * from Win32_OperatingSystem")
Set getBIOSSettings = getWMI_obj.ExecQuery ("Select * from Win32_BIOS")
Set IPConfigSet     = getWMI_obj.ExecQuery ("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
Set colSettingsCPU  = getWMI_obj.ExecQuery ("select ProcessorId, MaxClockSpeed from Win32_Processor")
Set colSettingsGPU  = getWMI_obj.ExecQuery ("select Caption, Description, DeviceName from Win32_DisplayConfiguration")
Set colSettingsVDU  = getWMI_obj.ExecQuery ("select Caption, Description, DeviceID, DisplayType, MonitorManufacturer, MonitorType, Name, PNPDeviceID, ScreenHeight, ScreenWidth from Win32_DesktopMonitor")

Set getCompSettingsProd = getWMI_obj.ExecQuery ("Select * from Win32_ComputerSystemProduct")

' # Get the UUID Details
' ##################################
'For Each objItem in getCompSettingsProd
'	msgbox objItem.caption & ", " & objItem.description & ", " & objItem.name & ", " & objItem.skunumber & ", " & objItem.vendor & ", " & objItem.version & ", " & objItem.uuid
'Next


' # Start the Main support Details
' ##################################

HeadlineInfo = "Main Support Details" & VbCr & VbCr & "****************************************" & vbCr & "Created by Jack Henry | https://github.com/MetalH47K/" & VbCr & VbCr & VbCr


For Each objItem in getCompSettings
	HeadlineInfo = HeadlineInfo & "Computer Name " & VbTab & ": " & objItem.Name & VbCr
	HeadlineInfo = HeadlineInfo & "User Name " & VbTab & ": " & objItem.UserName & VbCr
Next


ipNum = 0
For Each IPConfig in IPConfigSet
	ipNum = ipNum + 1
	If Not IsNull(IPConfig.IPAddress) Then
		For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
			HeadlineInfo = HeadlineInfo & "IP Address (" & ipNum & ")" & vbTab & ": " & IPConfig.IPAddress(i) & VbCr
		Next
	End If
Next


For Each objOS in getOSSettings
    dtmBootup = objOS.LastBootUpTime
    dtmLastBootUpTime = WMIDateStringToDate(dtmBootup)
    dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
Next

HeadlineInfo = HeadlineInfo & "Uptime (in hours) " & VbTab & ": " & dtmSystemUptime & VbCr



' # Start the Computer Details Code
' ##################################

getWMI_str = HeadlineInfo & vbCr & "Computer Details" & vbCr

For Each getComputer in getCompSettings
	getWMI_str = getWMI_str & _ 
	"   Manufacturer " & VbTab & ": " & getComputer.Manufacturer & VbCr & _ 
	"   Model " & VbTab & VbTab & ": " & getComputer.Model & VbCr & _ 
	"   Memory " & VbTab & ": " & round (getComputer.TotalPhysicalMemory / 1024 / 1024,0) & " MB" & VbCr
	CPUNo = 1
	For Each ObjCPU in colSettingsCPU
		getWMI_str = getWMI_str & "   Processor " & CPUNo & VbTab & ": " & objCPU.MaxClockSpeed & " Mhz" & VbCr
		CPUNo = CPUNo + 1	
	Next
Next


' # Start the OS Details Code
' ##################################

getWMI_str = getWMI_str & vbCr & "Operating System Details" & vbCr

For Each getComputer in getOSSettings
	getWMI_str = getWMI_str & _ 
	"   OS Version " & vbTab & ": " & getComputer.Caption & ", " & getComputer.CSDVersion & VbCr & _ 
	"   Version " & VbTab & ": " & getComputer.Version & VbCr & _ 
	"   Install Date " & VbTab & ": " & WMIDateStringToDate(getComputer.InstallDate) & VbCr & _ 
	"   Windows Folder" & vbTab & ": " & getComputer.WindowsDirectory & VbCr
Next

' # Start the Graphics Details Code
' ##################################

getWMI_str = getWMI_str & vbCr & "Graphics Card Details" & vbCr

For Each ObjGPU in colSettingsGPU
	getWMI_str = getWMI_str & _ 
	"   Graphics Card " & vbTab & ": " & objGPU.Description & VbCr
Next

monNum = 1
For Each ObjVDU in colSettingsVDU
	If objVDU.MonitorManufacturer <> "" Then
		getWMI_str = getWMI_str & VbCr & _
		"   Monitor " & monNum & VbCr & _
		"      Make " & vbTab & ": " & objVDU.MonitorManufacturer & VbCr & _
		"      Description" & vbTab & ": " & objVDU.Description & VbCr & _
		"      Resolution" & vbTab & ": " & objVDU.ScreenWidth & " x " & objVDU.ScreenHeight & VbCr
		monNum = monNum + 1
	End If
Next


' # Start the BIOS Details Code
' ##################################

getWMI_str = getWMI_str & vbCr & "BIOS Details" & vbCr

For Each getComputer in getBIOSSettings
	getWMI_str = getWMI_str & _ 
	"   Serial Number " & VbTab & ": " & getComputer.SerialNumber & VbCr & _ 
	"   Manufacturer " & VbTab & ": " & getComputer.Manufacturer & VbCr & _ 
	"   Version " & VbTab & ": " & getComputer.Version
Next

MsgBox getWMI_str,64,"Computer Support Information"

Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate =  _
        CDate(Mid(dtmBootup, 5, 2) & "/" & _
        Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
        & " " & Mid (dtmBootup, 9, 2) & ":" & _
        Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, 13, 2))
End Function

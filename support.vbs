'  VB Script to display computer details

On Error Resume Next

' Define Local Variables
' ######################

Dim objWMI
Dim colSettingsComp
Dim colSettingsOS
Dim colSettingsBios
Dim objComputer
Dim strWMI
Dim HeadlineInfo
Dim ipNo
Dim MonNo

Const vbLongDate = 1


' # Define queries
' #################
'Set objWMI = GetObject("winmgmts:")

Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Set colSettingsComp = objWMI.ExecQuery ("Select * from Win32_ComputerSystem")
Set colSettingsOS   = objWMI.ExecQuery ("Select * from Win32_OperatingSystem")
Set colSettingsBios = objWMI.ExecQuery ("Select * from Win32_BIOS")
Set IPConfigSet     = objWMI.ExecQuery ("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
Set colSettingsCPU  = objWMI.ExecQuery ("select ProcessorId, MaxClockSpeed from Win32_Processor")
Set colSettingsGPU  = objWMI.ExecQuery ("select Caption, Description, DeviceName from Win32_DisplayConfiguration")
Set colSettingsVDU  = objWMI.ExecQuery ("select Caption, Description, DeviceID, DisplayType, MonitorManufacturer, MonitorType, Name, PNPDeviceID, ScreenHeight, ScreenWidth from Win32_DesktopMonitor")

Set colSettingsCompProd = objWMI.ExecQuery ("Select * from Win32_ComputerSystemProduct")

' # Get the UUID Details
' ##################################
'For Each objItem in colSettingsCompProd
'	msgbox objItem.caption & ", " & objItem.description & ", " & objItem.name & ", " & objItem.skunumber & ", " & objItem.vendor & ", " & objItem.version & ", " & objItem.uuid
'Next


' # Start the Main support Details
' ##################################

HeadlineInfo = "Main Support Details" & VbCr & "********************" & VbCr


For Each objItem in colSettingsComp
	HeadlineInfo = HeadlineInfo & "Computer Name " & VbTab & ": " & objItem.Name & VbCr
	HeadlineInfo = HeadlineInfo & "User Name " & VbTab & ": " & objItem.UserName & VbCr
Next


ipNo = 0
For Each IPConfig in IPConfigSet
	ipNo = ipNo + 1
	If Not IsNull(IPConfig.IPAddress) Then
		For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
			HeadlineInfo = HeadlineInfo & "IP Address (" & ipNo & ")" & vbTab & ": " & IPConfig.IPAddress(i) & VbCr
		Next
	End If
Next


For Each objOS in colSettingsOS
    dtmBootup = objOS.LastBootUpTime
    dtmLastBootUpTime = WMIDateStringToDate(dtmBootup)
    dtmSystemUptime = DateDiff("h", dtmLastBootUpTime, Now)
Next

HeadlineInfo = HeadlineInfo & "Uptime (in hours) " & VbTab & ": " & dtmSystemUptime & VbCr



' # Start the Computer Details Code
' ##################################

strWMI = HeadlineInfo & vbCr & "Computer Details" & vbCr

For Each objComputer in colSettingsComp
	strWMI = strWMI & _ 
	"   Manufacturer " & VbTab & ": " & objComputer.Manufacturer & VbCr & _ 
	"   Model " & VbTab & VbTab & ": " & objComputer.Model & VbCr & _ 
	"   Memory " & VbTab & ": " & round (objComputer.TotalPhysicalMemory / 1024 / 1024,0) & " MB" & VbCr
	CPUNo = 1
	For Each ObjCPU in colSettingsCPU
		strWMI = strWMI & "   Processor " & CPUNo & VbTab & ": " & objCPU.MaxClockSpeed & " Mhz" & VbCr
		CPUNo = CPUNo + 1	
	Next
Next


' # Start the OS Details Code
' ##################################

strWMI = strWMI & vbCr & "Operating System Details" & vbCr

For Each objComputer in colSettingsOS
	strWMI = strWMI & _ 
	"   Windows Version " & vbTab & ": " & objComputer.Caption & ", " & objComputer.CSDVersion & VbCr & _ 
	"   Version " & VbTab & VbTab & ": " & objComputer.Version & VbCr & _ 
	"   Install Date " & VbTab & ": " & WMIDateStringToDate(objComputer.InstallDate) & VbCr & _ 
	"   Windows Folder" & vbTab & ": " & objComputer.WindowsDirectory & VbCr
Next

' # Start the Graphics Details Code
' ##################################

strWMI = strWMI & vbCr & "Graphics Card Details" & vbCr

For Each ObjGPU in colSettingsGPU
	strWMI = strWMI & _ 
	"   Graphics Card " & vbTab & ": " & objGPU.Description & VbCr
Next

MonNo = 1
For Each ObjVDU in colSettingsVDU
	If objVDU.MonitorManufacturer <> "" Then
		strWMI = strWMI & VbCr & _
		"   Monitor " & MonNo & VbCr & _
		"      Make " & vbTab & vbTab & ": " & objVDU.MonitorManufacturer & VbCr & _
		"      Description" & vbTab & ": " & objVDU.Description & VbCr & _
		"      Resolution" & vbTab & ": " & objVDU.ScreenWidth & " x " & objVDU.ScreenHeight & VbCr
		MonNo = MonNo + 1
	End If
Next


' # Start the BIOS Details Code
' ##################################

strWMI = strWMI & vbCr & "BIOS Details" & vbCr

For Each objComputer in colSettingsBios
	strWMI = strWMI & _ 
	"   Serial Number " & VbTab & ": " & objComputer.SerialNumber & VbCr & _ 
	"   Manufacturer " & VbTab & ": " & objComputer.Manufacturer & VbCr & _ 
	"   Version " & VbTab & VbTab & ": " & objComputer.Version
Next

''wscript.echo strWMI

MsgBox strWMI,64,"Computer Support Information"

Function WMIDateStringToDate(dtmBootup)
    WMIDateStringToDate =  _
        CDate(Mid(dtmBootup, 5, 2) & "/" & _
        Mid(dtmBootup, 7, 2) & "/" & Left(dtmBootup, 4) _
        & " " & Mid (dtmBootup, 9, 2) & ":" & _
        Mid(dtmBootup, 11, 2) & ":" & Mid(dtmBootup, 13, 2))
End Function
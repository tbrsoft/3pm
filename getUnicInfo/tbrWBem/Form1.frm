VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "F2"
      Height          =   735
      Left            =   90
      TabIndex        =   11
      Top             =   1440
      Width           =   1605
   End
   Begin VB.TextBox Text3 
      Height          =   555
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "Form1.frx":0442
      Top             =   7020
      Width           =   6225
   End
   Begin VB.TextBox Text2 
      Height          =   1425
      Left            =   2310
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Form1.frx":0448
      Top             =   5520
      Width           =   6195
   End
   Begin VB.ListBox List2 
      Columns         =   3
      Height          =   1815
      Left            =   2310
      TabIndex        =   8
      Top             =   3660
      Width           =   6165
   End
   Begin VB.CommandButton Command6 
      Caption         =   "TODOS"
      Height          =   405
      Left            =   0
      TabIndex        =   7
      Top             =   7650
      Width           =   2145
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Pais e idioma"
      Height          =   405
      Left            =   90
      TabIndex        =   6
      Top             =   7080
      Width           =   2145
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Registro"
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   2145
   End
   Begin VB.CommandButton Command3 
      Caption         =   "GetSystemInfo"
      Height          =   405
      Left            =   90
      TabIndex        =   4
      Top             =   3630
      Width           =   2145
   End
   Begin VB.ListBox List1 
      Columns         =   2
      Height          =   2595
      Left            =   2460
      TabIndex        =   3
      Top             =   1020
      Width           =   6045
   End
   Begin VB.CommandButton Command2 
      Caption         =   "WBem"
      Height          =   405
      Left            =   90
      TabIndex        =   2
      Top             =   990
      Width           =   2145
   End
   Begin VB.TextBox Text1 
      Height          =   945
      Left            =   2430
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":044E
      Top             =   60
      Width           =   6195
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Memory"
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   2145
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Text1 = GetBIOSDate
End Sub

Private Sub Command2_Click()
    List1.Clear
    On Error GoTo MUestraERR
    
    Dim ObjSet As SWbemObjectSet
    Dim SERV As SWbemServices
    
    Set SERV = GetObject("WinMgmts:")
    
    Set ObjSet = Nothing
    'datos del ventilador
    List1.AddItem "-----------------------"
    List1.AddItem "INFORMACION DE LA BIOS"
    Set ObjSet = SERV.InstancesOf("Win32_Bios")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each BIOS In ObjSet
            List1.AddItem "BiosCharacteristics: " + CStr(BIOS.BiosCharacteristics)
            List1.AddItem "BuildNumber: " + CStr(BIOS.BuildNumber)
            List1.AddItem "Caption: " + CStr(BIOS.Caption)
            List1.AddItem "CodeSet: " + CStr(BIOS.CodeSet)
            List1.AddItem "CurrentLanguage: " + CStr(BIOS.CurrentLanguage)
            List1.AddItem "Description: " + CStr(BIOS.Description)
            List1.AddItem "IdentificationCode: " + CStr(BIOS.IdentificationCode)
            List1.AddItem "InstallableLanguages: " + CStr(BIOS.InstallableLanguages)
            List1.AddItem "InstallDate: " + CStr(BIOS.InstallDate)
            List1.AddItem "LanguageEdition: " + CStr(BIOS.LanguageEdition)
            List1.AddItem "ListOfLanguages: " + CStr(BIOS.ListOfLanguages)
            List1.AddItem "Manufacturer: " + CStr(BIOS.Manufacturer)
            List1.AddItem "Name: " + CStr(BIOS.Name)
            List1.AddItem "OtherTargetOS: " + CStr(BIOS.OtherTargetOS)
            List1.AddItem "PrimaryBIOS: " + CStr(BIOS.PrimaryBIOS)
            List1.AddItem "ReleaseDate: " + CStr(BIOS.ReleaseDate)
            List1.AddItem "SerialNumber: " + CStr(BIOS.SerialNumber)
            List1.AddItem "BuildNumber: " + CStr(BIOS.BuildNumber)
            List1.AddItem "SMBIOSBIOSVersion: " + CStr(BIOS.SMBIOSBIOSVersion)
            List1.AddItem "SMBIOSMajorVersion: " + CStr(BIOS.SMBIOSMajorVersion)
            List1.AddItem "SMBIOSMinorVersion: " + CStr(BIOS.SMBIOSMinorVersion)
            List1.AddItem "SMBIOSPresent: " + CStr(BIOS.SMBIOSPresent)
            List1.AddItem "SoftwareElementID: " + CStr(BIOS.SoftwareElementID)
            List1.AddItem "SoftwareElementState: " + CStr(BIOS.SoftwareElementState)
            List1.AddItem "Status: " + CStr(BIOS.Status)
            List1.AddItem "TargetOperatingSystem: " + CStr(BIOS.TargetOperatingSystem)
            List1.AddItem "Version: " + CStr(BIOS.Version)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    
    Set ObjSet = Nothing
    'Win32_Processor Represents a device capable of interpreting a sequence of machine instructions on a Win32 computer system.
    'datos del ventilador
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION DEL PROCESADOR"
    Set ObjSet = SERV.InstancesOf("Win32_Processor")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each MICRO In ObjSet
            List1.AddItem "Availability: " + CStr(MICRO.Availability)
            List1.AddItem "AddressWidth: " + CStr(MICRO.AddressWidth)
            List1.AddItem "Architecture: " + CStr(MICRO.Architecture)
            List1.AddItem "CpuStatus: " + CStr(MICRO.CpuStatus)
            List1.AddItem "CreationClassName: " + CStr(MICRO.CreationClassName)
            List1.AddItem "CurrentClockSpeed: " + CStr(MICRO.CurrentClockSpeed)
            List1.AddItem "CurrentVoltage: " + CStr(MICRO.CurrentVoltage)
            List1.AddItem "DataWidth: " + CStr(MICRO.DataWidth)
            List1.AddItem "Description: " + CStr(MICRO.Description)
            List1.AddItem "DeviceID: " + CStr(MICRO.DeviceID)
            List1.AddItem "ExtClock: " + CStr(MICRO.ExtClock)
            List1.AddItem "Family: " + CStr(MICRO.Family)
            List1.AddItem "L2CacheSize: " + CStr(MICRO.L2CacheSize)
            List1.AddItem "L2CacheSpeed: " + CStr(MICRO.L2CacheSpeed)
            List1.AddItem "Level: " + CStr(MICRO.Level)
            List1.AddItem "LoadPercentage: " + CStr(MICRO.LoadPercentage)
            List1.AddItem "Manufacturer: " + CStr(MICRO.Manufacturer)
            List1.AddItem "MaxClockSpeed: " + CStr(MICRO.MaxClockSpeed)
            List1.AddItem "Name: " + CStr(MICRO.Name)
            List1.AddItem "OtherFamilyDescription: " + CStr(MICRO.OtherFamilyDescription)
            List1.AddItem "PNPDeviceID: " + CStr(MICRO.PNPDeviceID)
            List1.AddItem "ProcessorId: " + CStr(MICRO.ProcessorId)
            List1.AddItem "ProcessorType: " + CStr(MICRO.ProcessorType)
            List1.AddItem "Revision: " + CStr(MICRO.Revision)
            List1.AddItem "Role: " + CStr(MICRO.Role)
            List1.AddItem "SocketDesignation: " + CStr(MICRO.SocketDesignation)
            List1.AddItem "Status: " + CStr(MICRO.Status)
            List1.AddItem "StatusInfo: " + CStr(MICRO.StatusInfo)
            List1.AddItem "Stepping: " + CStr(MICRO.Stepping)
            List1.AddItem "SystemCreationClassName: " + CStr(MICRO.SystemCreationClassName)
            List1.AddItem "SystemName: " + CStr(MICRO.SystemName)
            List1.AddItem "UniqueId: " + CStr(MICRO.UniqueId)
            List1.AddItem "UpgradeMethod: " + CStr(MICRO.UpgradeMethod)
            List1.AddItem "Version: " + CStr(MICRO.Version)
            List1.AddItem "VoltageCaps: " + CStr(MICRO.VoltageCaps)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    Set ObjSet = Nothing
    'Win32_OnBoardDevice  Represents common adapter devices built into the motherboard (system board).
    'datos de DISPOSITIVOS ON BOARD
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION ONBOARD"
    Set ObjSet = SERV.InstancesOf("Win32_OnBoardDevice")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each ONBOARD In ObjSet
            List1.AddItem "Caption: " + CStr(ONBOARD.Caption)
            List1.AddItem "CreationClassName: " + CStr(ONBOARD.CreationClassName)
            List1.AddItem "Description: " + CStr(ONBOARD.Description)
            List1.AddItem "DeviceType: " + CStr(ONBOARD.DeviceType)
            List1.AddItem "Enabled: " + CStr(ONBOARD.Enabled)
            List1.AddItem "HotSwappable: " + CStr(ONBOARD.HotSwappable)
            List1.AddItem "InstallDate: " + CStr(ONBOARD.InstallDate)
            List1.AddItem "Manufacturer: " + CStr(ONBOARD.Manufacturer)
            List1.AddItem "Model: " + CStr(ONBOARD.Model)
            List1.AddItem "Name: " + CStr(ONBOARD.Name)
            List1.AddItem "OtherIdentifyingInfo: " + CStr(ONBOARD.OtherIdentifyingInfo)
            List1.AddItem "PartNumber: " + CStr(ONBOARD.PartNumber)
            List1.AddItem "PoweredOn: " + CStr(ONBOARD.PoweredOn)
            List1.AddItem "Removable: " + CStr(ONBOARD.Removable)
            List1.AddItem "Replaceable: " + CStr(ONBOARD.Replaceable)
            List1.AddItem "SerialNumber: " + CStr(ONBOARD.SerialNumber)
            List1.AddItem "SKU: " + CStr(ONBOARD.SKU)
            List1.AddItem "Status: " + CStr(ONBOARD.Status)
            List1.AddItem "Tag: " + CStr(ONBOARD.Tag)
            List1.AddItem "Version: " + CStr(ONBOARD.Version)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    Set ObjSet = Nothing
    'Win32_MotherboardDevice Represents a device that contains the central components of the Win32 computer system.
    'datos de DISPOSITIVOS ON BOARD
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION MOTHR DEVICE"
    Set ObjSet = SERV.InstancesOf("Win32_MotherboardDevice")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
        For Each ONBOARD In ObjSet
            List1.AddItem "Availability: " + CStr(ONBOARD.Availability)
            List1.AddItem "Caption: " + CStr(ONBOARD.Caption)
            List1.AddItem "Description: " + CStr(ONBOARD.Description)
            List1.AddItem "DeviceID: " + CStr(ONBOARD.DeviceID)
            List1.AddItem "InstallDate: " + CStr(ONBOARD.InstallDate)
            List1.AddItem "Name: " + CStr(ONBOARD.Name)
            List1.AddItem "PNPDeviceID: " + CStr(ONBOARD.PNPDeviceID)
            List1.AddItem "PowerManagementCapabilities: " + CStr(ONBOARD.PowerManagementCapabilities)
            List1.AddItem "PowerManagementSupported: " + CStr(ONBOARD.PowerManagementSupported)
            List1.AddItem "PrimaryBusType: " + CStr(ONBOARD.PrimaryBusType)
            List1.AddItem "RevisionNumber: " + CStr(ONBOARD.RevisionNumber)
            List1.AddItem "SecondaryBusType: " + CStr(ONBOARD.SecondaryBusType)
            List1.AddItem "Status: " + CStr(ONBOARD.Status)
            List1.AddItem "StatusInfo: " + CStr(ONBOARD.StatusInfo)
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    
    Set ObjSet = Nothing
    'Win32_SystemSlot Represents physical connection points including ports, motherboard slots and peripherals, and proprietary connections points.
    List1.AddItem "--------------------------"
    List1.AddItem "INFORMACION SLOTS"
    
    Set ObjSet = SERV.InstancesOf("Win32_SystemSlot")
    If ObjSet.Count > 0 Then
        List1.AddItem "Se encontaron: " + CStr(ObjSet.Count)
    
        For Each slots In ObjSet
            List1.AddItem "Caption: " + CStr(slots.Caption)
            List1.AddItem "ConnectorPinout: " + CStr(slots.ConnectorPinout)
            List1.AddItem "ConnectorType: " + CStr(slots.ConnectorType)
            List1.AddItem "CurrentUsage: " + CStr(slots.CurrentUsage)
            List1.AddItem "Description: " + CStr(slots.Description)
            List1.AddItem "HeightAllowed: " + CStr(slots.HeightAllowed)
            List1.AddItem "InstallDate: " + CStr(slots.InstallDate)
            List1.AddItem "LengthAllowed: " + CStr(slots.LengthAllowed)
            List1.AddItem "Manufacturer: " + CStr(slots.Manufacturer)
            List1.AddItem "MaxDataWidth: " + CStr(slots.MaxDataWidth)
            List1.AddItem "Model: " + CStr(slots.Model)
            List1.AddItem "Name: " + CStr(slots.Name)
            List1.AddItem "Number: " + CStr(slots.Number)
            List1.AddItem "OtherIdentifyingInfo: " + CStr(slots.OtherIdentifyingInfo)
            List1.AddItem "PartNumber: " + CStr(slots.PartNumber)
            List1.AddItem "PMESignal: " + CStr(slots.PMESignal)
            List1.AddItem "PoweredOn: " + CStr(slots.PoweredOn)
            List1.AddItem "PurposeDescription: " + CStr(slots.PurposeDescription)
            List1.AddItem "SerialNumber: " + CStr(slots.SerialNumber)
            List1.AddItem "Shared: " + CStr(slots.Shared)
            List1.AddItem "SKU: " + CStr(slots.SKU)
            List1.AddItem "SlotDesignation: " + CStr(slots.SlotDesignation)
            List1.AddItem "SpecialPurpose: " + CStr(slots.SpecialPurpose)
            List1.AddItem "Status: " + CStr(slots.Status)
            List1.AddItem "SupportsHotPlug: " + CStr(slots.SupportsHotPlug)
            List1.AddItem "Tag: " + CStr(slots.Tag)
            List1.AddItem "ThermalRating: " + CStr(slots.ThermalRating)
            List1.AddItem "VccMixedVoltageSupport: " + CStr(slots.VccMixedVoltageSupport)
            List1.AddItem "Version: " + CStr(slots.Version)
            List1.AddItem "VppMixedVoltageSupport: " + CStr(slots.VppMixedVoltageSupport)
        
        Next
    Else
        List1.AddItem "No se encontraron"
    End If
    Exit Sub
MUestraERR:
    Resume Next
End Sub

Private Sub Command3_Click()
    List2.Clear
    Dim INFO As SYSTEM_INFO
    GetSystemInfo INFO
    
    List2.AddItem "dwActiveProcessorMask: " + CStr(INFO.dwActiveProcessorMask)
    List2.AddItem "dwAllocationGranularity: " + CStr(INFO.dwAllocationGranularity)
    List2.AddItem "dwNumberOrfProcessors: " + CStr(INFO.dwNumberOrfProcessors)
    List2.AddItem "dwOemID: " + CStr(INFO.dwOemID)
    List2.AddItem "dwPageSize: " + CStr(INFO.dwPageSize)
    List2.AddItem "dwProcessorType: " + CStr(INFO.dwProcessorType)
    'joya este es unico en cada PC (todos los amd son iguales!!!)
    List2.AddItem "dwReserved: " + Str(INFO.dwReserved)
    'joya este es unico en cada PC
    List2.AddItem "lpMaximumApplicationAddress: " + CStr(INFO.lpMaximumApplicationAddress)
    List2.AddItem "lpMinimumApplicationAddress: " + CStr(INFO.lpMinimumApplicationAddress)
End Sub

Private Sub Command4_Click()
    Dim TodoRet As String
    Dim Ret As String
    
    Ret = GetString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "Identifier")
    Ret = "Identifier: " + Ret
    TodoRet = Ret
    
    Ret = GetString(HKEY_LOCAL_MACHINE, "Hardware\Description\System\CentralProcessor\0", "VendorIdentifier")
    Ret = "Vendor Identifier: " + Ret
    TodoRet = TodoRet + vbCrLf + Ret
    
    Ret = GetString(HKEY_LOCAL_MACHINE, "Enum\root\*pnp0c01\0000", "CPU")
    Ret = "CPU: " + Ret
    TodoRet = TodoRet + vbCrLf + Ret
    
    Ret = GetString(HKEY_LOCAL_MACHINE, "Enum\root\*pnp0c01\0000", "BiosName")
    Ret = "Bios Name: " + Ret
    TodoRet = TodoRet + vbCrLf + Ret
    
    Ret = GetString(HKEY_LOCAL_MACHINE, "Enum\root\*pnp0c01\0000", "BiosVersion")
    Ret = "Bios Version: " + Ret
    TodoRet = TodoRet + vbCrLf + Ret
    
    Ret = GetString(HKEY_LOCAL_MACHINE, "Enum\root\*pnp0c01\0000", "BiosDate")
    Ret = "Bios Date: " + Ret
    TodoRet = TodoRet + vbCrLf + Ret
    Text2 = TodoRet

End Sub

Private Sub Command5_Click()
    Text3 = "Pais: " & GetInfo(LOCALE_SENGCOUNTRY) + _
        " (" + GetInfo(LOCALE_SNATIVECTRYNAME) & ")," & _
        vbCrLf & "Idioma: " & _
        GetInfo(LOCALE_SENGLANGUAGE) & " (" & _
        GetInfo(LOCALE_SNATIVELANGNAME) + ")"
End Sub

Private Sub Command6_Click()
    Call Command1_Click
    Call Command2_Click
    Call Command3_Click
    Call Command4_Click
    Call Command5_Click
End Sub

Private Sub Command7_Click()
    Form2.Show 1
End Sub

Private Sub Form_Load()
    Me.Caption = "GetUnicInfo - tbrGUI v" + CStr(App.Major) + "." + CStr(App.Revision)
End Sub


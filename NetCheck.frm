VERSION 5.00
Begin VB.Form frmNetCheck 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Network Check"
   ClientHeight    =   4965
   ClientLeft      =   435
   ClientTop       =   1950
   ClientWidth     =   5460
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   5460
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   5280
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3015
      TabIndex        =   1
      Top             =   4455
      Width           =   1185
   End
   Begin VB.CommandButton cmdNetCheck 
      Caption         =   "NetCheck"
      Default         =   -1  'True
      Height          =   375
      Left            =   1530
      TabIndex        =   0
      Top             =   4455
      Width           =   1185
   End
End
Attribute VB_Name = "frmNetCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const HKEY_LOCAL_MACHINE       As Long = &H80000002
Dim m_clsRegAccess                     As clsRegistryAccess

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdNetCheck_Click()
   Dim p_vntRtn As Variant
   Dim p_vntAdapters As Variant
   Dim p_strSubKey As String
   Dim p_strValueName As String
   Dim p_lngNumAdapters As Long
   Dim p_lngLoop As Long
   Dim p_lngPos As Long
   Dim p_strAdapterName As String
   Dim p_strTmp As String
   Dim p_blnFirstTime As Boolean
   
   ' ------------------------------------------
   ' Clear the lise
   ' ------------------------------------------
   Me.List1.Clear
   
   ' ------------------------------------------
   ' Get the number of adapters in system
   ' ------------------------------------------
   p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters"
   p_vntAdapters = m_clsRegAccess.EnumerateRegistryKeys(p_strSubKey, HKEY_LOCAL_MACHINE)
   
   On Error Resume Next
   p_lngNumAdapters = UBound(p_vntAdapters)
   If p_lngNumAdapters <= 0 Then
      ' ---------------------------------------
      ' Exit if no adapters found
      ' ---------------------------------------
      Me.List1.AddItem "No network adapters found"
      Exit Sub
   End If
   On Error GoTo 0
   
   ' ------------------------------------------
   ' Get the TCPIP parameters: Domain
   ' ------------------------------------------
   p_strSubKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
   p_strValueName = "Domain"
   p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
   Me.List1.AddItem "Domain: " & CStr(p_vntRtn)
   
   ' ------------------------------------------
   ' Get the TCPIP parameters: Host Name
   ' ------------------------------------------
   p_strValueName = "HostName"
   p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
   Me.List1.AddItem "Host Name: " & CStr(p_vntRtn)
   
   ' ------------------------------------------
   ' Get the TCPIP parameters: one or more name servers
   ' ------------------------------------------
   p_strValueName = "NameServer"
   p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
   p_strTmp = CStr(p_vntRtn)
   p_blnFirstTime = True
   p_lngPos = InStr(1, p_strTmp, " ", vbTextCompare)
   Do While p_lngPos > 0
      If p_blnFirstTime = True Then
         Me.List1.AddItem "Name Server(s): " & Trim$(Mid$(p_strTmp, 1, p_lngPos - 1))
         p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
         p_blnFirstTime = False
      Else
         Me.List1.AddItem Space$(6) & Trim$(Mid$(p_strTmp, 1, p_lngPos - 1))
         p_strTmp = Mid$(p_strTmp, p_lngPos + 1)
      End If
      p_lngPos = InStr(1, p_strTmp, " ", vbTextCompare)
   Loop
   If Len(p_strTmp) > 0 Then
      Me.List1.AddItem Space$(6) & Trim$(p_strTmp)
   End If
   Me.List1.AddItem "--------------------------------------"
   
   ' ------------------------------------------
   ' Get info for each adapter
   ' ------------------------------------------
   For p_lngLoop = 1 To p_lngNumAdapters
      If p_lngLoop > 1 Then
         Me.List1.AddItem "--------------------------------------"
      End If
      
      p_strAdapterName = p_vntAdapters(p_lngLoop)(0)
      Me.List1.AddItem "Adapter Name: " & p_strAdapterName
   
      ' ---------------------------------------
      ' Name server -- IP Address of primary WINS server
      ' ---------------------------------------
      p_strSubKey = "SYSTEM\CurrentControlSet\Services\NetBT\Adapters\" & p_strAdapterName
      p_strValueName = "NameServer"
      p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
      Me.List1.AddItem "   Primary WINS Server: " & CStr(p_vntRtn)
   
      ' ---------------------------------------
      ' Backup name server
      ' ---------------------------------------
      p_strValueName = "NameServerBackup"
      p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
      Me.List1.AddItem "   WINS Server Backup: " & CStr(p_vntRtn)
      
      ' ---------------------------------------
      ' Default gateway
      ' ---------------------------------------
      p_strSubKey = "SYSTEM\CurrentControlSet\Services\" & p_strAdapterName & "\Parameters\Tcpip"
      p_strValueName = "DefaultGateway"
      p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
      Me.List1.AddItem "   Default Gateway(s): " & CStr(p_vntRtn)

      ' ---------------------------------------
      ' Is DHCP enabled?
      ' ---------------------------------------
      p_strValueName = "EnableDHCP"
      p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
      Me.List1.AddItem "   DHCP Enabled: " & CBool(p_vntRtn)

      ' ---------------------------------------
      ' IP Address
      ' ---------------------------------------
      p_strValueName = "IPAddress"
      p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
      Me.List1.AddItem "   IP Address: " & CStr(p_vntRtn)

      ' ---------------------------------------
      ' Get subnet mask
      ' ---------------------------------------
      p_strValueName = "SubnetMask"
      p_vntRtn = m_clsRegAccess.GetRegistryValue(p_strSubKey, p_strValueName, HKEY_LOCAL_MACHINE)
      Me.List1.AddItem "   Subnet Mask(s): " & CStr(p_vntRtn)

   Next p_lngLoop
   
End Sub

Private Sub Form_Load()
   ' ------------------------------------------
   ' Setup the class
   ' ------------------------------------------
   Set m_clsRegAccess = New clsRegistryAccess
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ' ------------------------------------------
   ' Clear out the class
   ' ------------------------------------------
   Set m_clsRegAccess = Nothing
End Sub

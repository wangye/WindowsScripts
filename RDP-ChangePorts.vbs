'----------------------------------------------------------------------
' *********************************************************************
' * Abstract:
' *    Change default Windows Remote Desktop Port 3389 (RDP-TCP, TDS-TCP)
' *
' * Author  :
' *    Wang Ye <pcn88 at hotmail.com>
' * Date    :
' *    2012-01-25 00:11
' * Website :
' *    http://wangye.org/
' *
' * For more information please visit http://wangye.org/
' *
' *  Copyright notice
' *
' *  (c) 2008-2011 WANG Ye (pcn88 at hotmail.com)
' *  All rights reserved
' *
' *  This script is free software; you can redistribute it and/or modify
' *  it under the terms of the GNU General Public License as published by
' *  the Free Software Foundation; either version 2 of the License, or
' *  (at your option) any later version.
' *
' *  The GNU General Public License can be found at
' *  http://www.gnu.org/copyleft/gpl.html.
' *  A copy is found in the textfile GPL.txt and important notices to the license
' *  from the author is found in LICENSE.txt distributed with these scripts.
' *
' *  This script is distributed in the hope that it will be useful,
' *  but WITHOUT ANY WARRANTY; without even the implied warranty of
' *  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' *  GNU General Public License for more details.
' *
' *  This copyright notice MUST APPEAR in all copies of the script!
' *********************************************************************
'----------------------------------------------------------------------
' // http://www.microsoft.com/china/technet/community/scriptcenter/registry/default.mspx
Option Explicit

Const AppTitle = "Modify RDP Port Number"

Const StatusOk = 0
Const StatusInvalidPortNumber = -1
Const StatusSetRDPPortNumberFailed = -2
Const StatusSetTDSPortNumberFailed = -3

Const L_Invalid_PortNumber_Text = "ERROR : Invalid port number."
Const L_User_Cancelled_Text = "User cancelled."
Const HKEY_LOCAL_MACHINE = &H80000002
Const RDPTcpPath =_
 "SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\"
Const TDSTcpPath =_
 "SYSTEM\CurrentControlSet\Control\Terminal Server\Wds\rdpwd\Tds\tcp\"
 
Class RDPTS

  Private strComputer
  Private Registry
  
  Private Sub Class_Initialize()
    strComputer = "."
    Set Registry = GetObject(_
		"winmgmts:{impersonationLevel=impersonate}!\\" &_
      strComputer & "\root\default:StdRegProv")
  End Sub

  Private Sub Class_Terminate()
    Set Registry = Nothing
  End Sub
  
  Function isPortAlreadyExists(portnum)
    ' ÅÐ¶Ï¶Ë¿ÚÊÇ·ñ³åÍ»£¨ÉÐÎ´ÊµÏÖ£©
    isPortAlreadyExists = False
  End Function

  Public Function getPortNumber(lowerbound, upperbound)
    If lowerbound < 4 Then lowerbound = 4
    If upperbound > 65534 Then upperbound = 65535
    Do
      Randomize
      getPortNumber = Int(_
      (upperbound - lowerbound + 1)_
      * Rnd + lowerbound)
    Loop Until getPortNumber<>3389_
        And (Not isPortAlreadyExists(getPortNumber))
  End Function

  Public Function isPortValid(portnum)
    isPortValid = False
    If Not IsNumeric(portnum) Then
      Exit Function
    End If
    If portnum < 4 Then
      Exit Function
    End If
    If portnum > 65534 Then
      Exit Function
    End If
    isPortValid = True
  End Function
  
  Public Function getRDPTcpPortNumber()
    Registry.GetDWORDValue HKEY_LOCAL_MACHINE,_
		RDPTcpPath,"PortNumber",getRDPTcpPortNumber
  End Function
  
  Public Function getTDSTcpPortNumber()
    Registry.GetDWORDValue HKEY_LOCAL_MACHINE,_
		TDSTcpPath,"PortNumber",getTDSTcpPortNumber
  End Function
  
  Public Function setRDPTcpPortNumber(portnum)
    On Error Resume Next
    setRDPTcpPortNumber = True
    Registry.SetDWORDValue HKEY_LOCAL_MACHINE,_
		RDPTcpPath,"PortNumber",portnum
    If Err Then Err.Clear : setRDPTcpPortNumber = False
  End Function
  
  Public Function setTDSTcpPortNumber(portnum)
    On Error Resume Next
    setTDSTcpPortNumber = True
    Registry.SetDWORDValue HKEY_LOCAL_MACHINE,_
		TDSTcpPath,"PortNumber",portnum
    If Err Then Err.Clear : setTDSTcpPortNumber = False
  End Function
  
  Public Function addFirewallPolicy(portnum, name, state)
    Dim netfw, policy, port, ports
    Set netfw = WScript.CreateObject("HNetCfg.FwMgr")
    Set policy = netfw.LocalPolicy.CurrentProfile
    Set port = WScript.CreateObject("HNetCfg.FwOpenPort")
      port.Port = portnum
      port.Name = name
      port.Enabled = state
    Set ports = policy.GloballyOpenPorts
      addFirewallPolicy = ports.Add(port)
    Set ports = Nothing
    Set port = Nothing
    Set policy = Nothing
    Set netfw = Nothing
  End Function
End Class


Function VBMain()
  
  VBMain = StatusOk
  
  Dim RDS, portnum, source
  Set RDS = New RDPTS
  portnum = RDS.getPortNumber(3390, 65530)
  source = InputBox("Original port number detected:" & vbCrLf &_
          "RDP-TCP(" & RDS.getRDPTcpPortNumber() &_
          "), TDS-TCP(" & RDS.getTDSTcpPortNumber() &_
          ")" & vbCrLf & vbCrLf &_
          "Please Enter the new port number" & vbCrLf &_
          "for RDP(Terminal Services) Server", _
          AppTitle, portnum)
          
  If source = "" Then
    WScript.Echo L_User_Cancelled_Text
    Exit Function
  End If
  
  If Not RDS.isPortValid(source) Then
    WScript.Echo L_Invalid_PortNumber_Text
    VBMain = StatusInvalidPortNumber
    Exit Function
  End If
  
  portnum = source
  
  If MsgBox("Pending changes : " & vbCrLf &  vbCrLf &_
       "RDP-TCP ` " & RDS.getRDPTcpPortNumber() &_
	   " -> " & portnum & " `" & vbCrLf &_
       "TDS-TCP ` " & RDS.getTDSTcpPortNumber() &_
	   " -> " & portnum & " `" & vbCrLf &  vbCrLf &_
	   "Are you sure?", vbOKCancel, AppTitle) = vbCancel Then
    WSH.Echo "Cancelled, No changes occured."
    Exit Function
  End If
  
  If Not RDS.setRDPTcpPortNumber(portnum) Then
    WSH.Echo "Set RDP-TCP port number `" &_
    RDS.getRDPTcpPortNumber() & "` to `" &_
    portnum & "` failed!"
    VBMain = StatusSetRDPPortNumberFailed
    Exit Function
  End If
  
  If Not RDS.setTDSTcpPortNumber(portnum) Then
    WSH.Echo "Set TDS-TCP port number `" &_
    RDS.getTDSTcpPortNumber() & "` to `" &_
    portnum & "` failed!"
    VBMain = StatusSetTDSPortNumberFailed
    Exit Function
  End If
  
  If MsgBox("Do you want add port `" & portnum &_
  "` to Windows Firewall policy?", vbOKCancel, AppTitle) = vbOK Then
    Do
      source = InputBox("Enter the name for this new policy",_
		AppTitle, "RDP(Terminal Services)")
      If source="" Then
        If MsgBox("Policy name required, Do you want quit Add Policy?",_
        vbOKCancel, AppTitle) = VbOK Then
          Exit Do
        End If
      End If
    Loop Until source<>""
    
    If source<>"" Then
      RDS.addFirewallPolicy portnum, source, 1
    End If
  End If
  WScript.Echo "All done successfully!" & vbCrLf &_
				"For more information please visit http://wangye.org/"
  Set RDS = Nothing
End Function

Call WScript.Quit(VBMain())

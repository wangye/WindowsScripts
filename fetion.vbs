'----------------------------------------------------------------------
' *********************************************************************
' * Abstract:
' *    Send Fetion Message via Fetion WAP website
' *
' * Usage:
' *      fetion  [-u account] [-p password]
' *              [-r receiver] [-m message]
' *              [--send=enable|disable]
' *              [--type=SMS|default]
' *              [--login=enable|disable]
' *              [--logout=enable|disable]
' *              [--echo=enable|disable]
' * Parameters:
' *   -u       - Sender account's phone number
' *   -p       - Sender account's password
' *   -r       - Receiver phone number
' *   -m       - Message text
' *   --send   - Use send features
' *   --type   - Use send type
' *   --login  - Login platform
' *   --logout - Logout platform
' *   --echo   - Echo on screen
' *
' * Author  :
' *    Wang Ye <pcn88 at hotmail.com>
' * Date    :
' *    2012-01-25 00:11
' * Website :
' *    http://wangye.org/
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
Option Explicit

Const L_Message_Login_Failed_Text = "-Login failed!"
Const L_Message_Login_Succeeded_Text = "+Login succeeded!"
Const L_Message_Logout_Failed_Text = "-Logout failed!"
Const L_Message_Logout_Succeeded_Text = "+Logout succeeded!"
Const L_Message_SendToOwn_Failed_Text = "-Send to '%1'(sender) failed!"
Const L_Message_SendToOwn_Succeeded_Text = "+Send to '%1'(sender) succeeded!"
Const L_Message_SendSMS_Failed_Text = "-Send SMS to '%1' failed!"
Const L_Message_SendSMS_Succeeded_Text = "+Send SMS to '%1' succeeded!"
Const L_Message_SendMsg_Failed_Text = "-Send to '%1' failed!"
Const L_Message_SendMsg_Succeeded_Text = "+Send to '%1' succeeded!"
Const L_Message_Display_OnScreen_Text = "Message : %1 (0x%2)"
Const L_Message_PhoneNumber_Invalid_Text = "-Phone number '%1' invalid!"
Const L_Message_PhoneNumber1_Invalid_Text = "-Phone number invalid!"
Const L_Message_PhonePassword_Invalid_Text = "-Phone password invalid!"
Const L_Message_MessageText_Invalid_Text = "-Message invalid!"
Const L_Message_Enter_SendNumber_Text = "Please enter the send number:"
Const L_Message_Enter_SendPassword_Text = "Please enter the send password:"
Const L_Message_Enter_MessageText_Text = "Please enter the message:"
Const L_Message_Option_Selection_Text = "Do you want continue [Y/n]?"
Const L_Argument_SendPhoneNumber_Name = "u"
Const L_Argument_SendPassword_Name = "p"
Const L_Argument_ReceivePhoneNumber_Name = "r"
Const L_Argument_MessageText_Name = "m"
Const L_Argument_HasLogin_Name = "login"
Const L_Argument_HasLogout_Name = "logout"
Const L_Argument_TypeName_Name = "type"
Const L_Argument_LoginStatus_Name = "status"
Const L_Argument_HasSend_Name = "send"
Const L_Argument_HasEcho_Name = "echo"
Const L_Argument_DisplayHelp_Name = "help"
Const L_Argument_Splitter_Token = "="

Const L_Help_Help_General01_Text = " Usage:"
Const L_Help_Help_General02_Text = "     fetion  [-u account] [-p password]"
Const L_Help_Help_General03_Text = "             [-r receiver] [-m message]"
Const L_Help_Help_General04_Text = "             [--send=enable|disable]"
Const L_Help_Help_General05_Text = "             [--type=SMS|default]"
Const L_Help_Help_General06_Text = "             [--login=enable|disable]"
Const L_Help_Help_General07_Text = "             [--logout=enable|disable]"
Const L_Help_Help_General08_Text = "             [--status=online|busy|leave|hidden]"
Const L_Help_Help_General09_Text = "             [--echo=enable|disable]"
const L_Help_Help_General10_Text = "Parameters:"
const L_Help_Help_General11_Text = "  -u       - Sender account's phone number"
const L_Help_Help_General12_Text = "  -p       - Sender account's password"
const L_Help_Help_General13_Text = "  -r       - Receiver phone number"
const L_Help_Help_General14_Text = "  -m       - Message text"
const L_Help_Help_General15_Text = "  --send   - Use send features"
const L_Help_Help_General16_Text = "  --type   - Use send type"
const L_Help_Help_General17_Text = "  --login  - Login platform"
const L_Help_Help_General18_Text = "  --logout - Logout platform"
const L_Help_Help_General19_Text = "  --status - Login status"
const L_Help_Help_General20_Text = "  --echo   - Echo on screen"

Dim EchoText

Class UriParser
    Private Function isUpper(c)
        isUpper = CBool( Asc(c) >= Asc("A") And Asc(c) <= Asc("Z") )
    End Function
    
    Private Function isLower(c)
        isLower = CBool( Asc(c) >= Asc("a") And Asc(c) <= Asc("z") )
    End Function
    
    Private Function isAlpha(c)
        isAlpha = CBool(isUpper(c) Or isLower(c))
    End Function
    
    Private Function isSpace(c)
        isSpace = CBool(Asc(c) = Asc(" "))
    End Function
    
    ' // Thanks Demon
    ' // See http://demon.tw/programming/vbs-php-urlencode.html
    Public Function encode(str, charset)
        Dim i,c
        For i = 1 To Len(str)
            c = Mid(str, i, 1)
            If isAlpha(c) Or c = "-" Or c = "_" Or c = "." Then
                encode = encode & c
            ElseIf isSpace(c) Then
                encode = encode & "+"
            Else
                If UCase(charset) = "UTF-8" Then
                   Dim s : s = c
                   c = "&H" & Hex(AscW(c))
                   If c >= &H0001 And c <= &H007F Then
                        encode = encode & s
                   ElseIf c > &H07FF Then
                        encode = encode & "%" & Hex(&HE0 Or (c\(2^12) And &H0F))
                        encode = encode & "%" & Hex(&H80 Or (c\(2^6) And &H3F))
                        encode = encode & "%" & Hex(&H80 Or (c\(2^0) And &H3F))
                    Else
                        encode = encode & "%" & Hex(&HC0 Or (c\(2^6) And &H1F))
                        encode = encode & "%" & Hex(&H80 Or (c\(2^0) And &H3F))
                    End If
                Else
                   c = Asc(c)
                   encode = encode & "%" & Left(Hex(c),2)
                   encode = encode & "%" & Right(Hex(c),2)
                End If
            End If
        Next
    End Function
End Class

' // Thanks Demon
' // See http://demon.tw/my-work/vbsfetion.html
Class FetionMessager
    Private BASE_URI
    Private INDEX_DIR
    Private LOGIN_SUBMIT_DIR
    Private LOGOUT_SUBMIT_DIR
    Private SMS_SUBMIT_DIR
    Private FETION_SUBMIT_DIR
    Private FETION_MYSELF_SUBMIT_DIR
    Private SEARCH_SUBMIT_DIR
    Private GPRS_WAP_IPADDR
    
    Private request_devid
    Private http
    Private content_charset
    Private regex
    
    Private status_code
    Private status_message
    Private error_occured
    
    Private Sub Class_Initialize()
        BASE_URI = "http://f.10086.cn"
        INDEX_DIR = "/im/index/indexcenter.action"
        LOGIN_SUBMIT_DIR = "/im/login/inputpasssubmit1.action"
        LOGOUT_SUBMIT_DIR = "/im/index/logoutsubmit.action"
        FETION_SUBMIT_DIR = "/im/chat/sendMsg.action"
        FETION_MYSELF_SUBMIT_DIR = "/im/user/sendMsgToMyselfs.action"
        SMS_SUBMIT_DIR = "/im/chat/sendShortMsg.action"
        SEARCH_SUBMIT_DIR = "/im/index/searchOtherInfoList.action"
        GPRS_WAP_IPADDR = "10.0.0.172"
        
        status_code = 0
        Set http = WSH.CreateObject("WinHttp.WinHttpRequest.5.1")
        Set regex = New RegExp
    End Sub
    
    Private Sub Class_Terminate()
        Set regex = Nothing
        Set http = Nothing
    End Sub
    
    Private Function encodeURI(str)
        Dim uri
        Set uri = New UriParser
        encodeURI = uri.encode(str, getContentCharset())
        Set uri = Nothing
    End Function
    
    Private Function buildLoginParameters(mobile, password, status)
        buildLoginParameters = "m=" & encodeURI(mobile) & "&pass=" & encodeURI(password) & "&loginstatus=" & status
    End Function
    
    Private Function buildMessageParameters(msg)
        buildMessageParameters = "msg=" & encodeURI(msg)
    End Function
    
    Private Function buildUserIdParameters(uid)
        buildUserIdParameters = "touserid=" & encodeURI(uid)
    End Function
    
    Private Function buildSearchTextParameters(text)
        buildSearchTextParameters = "searchText=" & encodeURI(text)
    End Function
    
    Private Function buildUrl(url, param)
        buildUrl = url & "?" & param
    End Function
    
    Private Function request(ByVal url, data, method)
        method = UCase(method)
        url = BASE_URI & url
        http.open method, url, False
        http.setRequestHeader "Accept", "text/xml,application/xml,application/xhtml+xml," & _
        "text/html;q=0.9,text/plain;q=0.8,text/vnd.wap.wml,image/png," &_
        "application/java-archive,application/java,application/x-java-archive," &_
        "text/vnd.sun.j2me.app-descriptor,application/vnd.oma.drm.message," &_
        "application/vnd.oma.drm.content,application/vnd.oma.dd+xml," &_
        "application/vnd.oma.drm.rights+xml,application/vnd.oma.drm.rights+wbxml," &_
        "application/x-nokia-widget,text/x-opml,*/*;q=0.5"
        
        http.setRequestHeader "User-Agent", "Mozilla/5.0 (SymbianOS/9.4; U; Series60/5.0 " &_
        "Nokia5800d-1/52.50.2008.37; Profile/MIDP-2.1 Configuration/CLDC-1.1 ) " &_
        "AppleWebKit/413 (KHTML, like Gecko) Safari/413"
        
        http.setRequestHeader "X-Forwarded-For", GPRS_WAP_IPADDR
        http.setRequestHeader "Forwarded-For", GPRS_WAP_IPADDR
        http.setRequestHeader "Client_IP", GPRS_WAP_IPADDR
        http.setRequestHeader "Client-IP", GPRS_WAP_IPADDR
        http.setRequestHeader "VIA", GPRS_WAP_IPADDR
        http.setRequestHeader "REMOTE_ADDR", GPRS_WAP_IPADDR
        http.setRequestHeader "REMOTE-ADDR", GPRS_WAP_IPADDR
        http.setRequestHeader "X-Nokia-MusicShop-Bearer", "GPRS/3G"
        http.setRequestHeader "X-Nokia-MusicShop-Version", "11.0842.9"
        http.setRequestHeader "X-Wap-Profile", "http://nds1.nds.nokia.com/uaprof/Nokia5800d-1r100-3G.xml"
        http.setRequestHeader "X-Online-Host", Replace(Replace(BASE_URI, "http://", ""), "https://", "")
        
        If Not IsEmpty(request_devid) Then
            http.setRequestHeader "x-up-calling-line-id", request_devid
            http.setRequestHeader "X-Up-subno", request_devid
        End If
        
        If method = "POST" Then
            http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            http.setRequestHeader "Referer", BASE_URI & INDEX_DIR
            http.send data
        Else
            http.send
        End If
        request = http.responseText
    End Function
    
    Private Function post(url, data)
        post = request(url, data, "POST")
    End Function
    
    Private Function getContentCharset()
        If IsEmpty(content_charset) Then
            Dim contentType
            Call request(INDEX_DIR, Null, "GET")
            contentType = http.getResponseHeader("Content-Type")
            regex.IgnoreCase = True
            regex.Global = False
            regex.Pattern = "^[\w\/\.]+; charset=(.+)$"
            If regex.Test(contentType) Then
               Dim matches
               Set matches = regex.Execute(contentType)
               content_charset = matches.Item(0).Submatches(0)
               Set matches = Nothing
            End If
        End If
        If IsEmpty(content_charset) Then content_charset = "UTF-8"
        getContentCharset = content_charset
    End Function
    
    Private Function parseLoginMessage(str)
        On Error Resume Next
        error_occured = False

        Dim login_status, login_message
        regex.IgnoreCase = True
        regex.Global = True
        regex.Pattern = "ontimer\=[""']" & _
            Replace(Replace(Replace(BASE_URI & INDEX_DIR, "/", "\/"), ".", "\."), ":", "\:")
        ' Succeeded
        If regex.Test(str) Then
            login_status = True
        Else
            regex.Pattern = "ontimer\=[""']" & _
                Replace(Replace(INDEX_DIR, "/", "\/"), ".", "\.")
            login_status = regex.Test(str)
            ' ontimer="/im/login/login.action
            ' login_status = False
        End If
        regex.Pattern = "timer value\=[""']\d+[""'] *\/>\s*<p>\s*(.+?[^<\s])\s*[^<]<br"
        If regex.Test(str) Then
            Dim matches
            Set matches = regex.Execute(str)
            login_message = matches.Item(0).Submatches(0)
            Set matches = Nothing
        End If
        If Err.Number<>0 Then
            login_status = Err.Number
            login_message = Err.Description
            error_occured = True
            Err.Clear
        End If
        parseLoginMessage = Array(login_status, login_message)
    End Function
    
    Private Function parseLogoutMessage(str)
        On Error Resume Next
        error_occured = False
        
        Dim logout_status, logout_message
        
        regex.IgnoreCase = True
        regex.Global = True
        regex.Pattern = "<card id\=""start"" title\="".+[^""]"">\s*<p>\s*<img"
        logout_status = regex.Test(str)
        If Not logout_status Then
            regex.Pattern = "<card id\=""start"" title\="".+[^""]"">\s*<p>\s*(.+?[^<\/\\\s])\s*<br"
            logout_status = regex.Test(str)
            If logout_status Then
                Dim matches
                Set matches = regex.Execute(str)
                logout_message = matches.Item(0).Submatches(0)
                Set matches = Nothing
            End If
        Else
            logout_status = False
        End If
        If Err.Number<>0 Then
            logout_status = Err.Number
            logout_message = Err.Description
            error_occured = True
            Err.Clear
        End If
        parseLogoutMessage = Array(logout_status, logout_message)
    End Function
    
    Private Function parseSendMessage(str)
        On Error Resume Next
        error_occured = False
        
        Dim send_status, send_message
        regex.IgnoreCase = True
        regex.Global = True
        regex.Pattern = "<card id\=""start"" title\=""(.+?[^""' ])"""
        send_status = regex.Test(str)
        If send_status Then
            Dim matches
            Set matches = regex.Execute(str)
            send_message = matches.Item(0).Submatches(0)
            ' send_message = "发送成功" Or send_message = "消息会话提示"
            ' fix Chinese characters in some platform
            If StrEqual(send_message, Array(21457, -28671, 25104, 21151)) Or _
               StrEqual(send_message, Array(28040, 24687, 20250, -29731, 25552, 31034)) Then
                send_status = True
            Else
                send_status = False
            End If
            Set matches = Nothing
            regex.Pattern = "timer value\=[""']\d+[""'] *\/>\s*<p>\s*(.+?[^<\s\/])\s*[^<]<br"
            If regex.Test(str) Then
                Set matches = regex.Execute(str)
                send_message = matches.Item(0).Submatches(0)
                Set matches = Nothing
            End If
        End If
        
        If Err.Number<>0 Then
            send_status = Err.Number
            send_message = Err.Description
            error_occured = True
            Err.Clear
        End If
        parseSendMessage = Array(send_status, send_message)
    End Function
    
    Private Function convertPhoneNumberToUserId(mobile)
        convertPhoneNumberToUserId = -1
        Dim content
        content = post(SEARCH_SUBMIT_DIR, buildSearchTextParameters(mobile))
        If http.status = 200 Then
            regex.IgnoreCase = True
            regex.Global = True
            regex.Pattern = "/toinputMsg\.action\?touserid=(\d+)"
            If regex.Test(content) Then
                Dim matches
                Set matches = regex.Execute(content)
                convertPhoneNumberToUserId = matches.Item(0).Submatches(0)
                Set matches = Nothing
            End If
        End If
    End Function
    
    Private Function StrEqual(str, arr)
        StrEqual = False
        If Len(str)<>(Ubound(arr)+1) Then
            Exit Function
        End If
        Dim i, n
        For i = 1 To Len(str)
            n = AscW(Mid(str, i, 1))
            If n<>arr(i-1) Then Exit Function
        Next
        StrEqual = True
    End Function
    
    Public Function isMobilePhoneNumberValid(mobile)
        regex.IgnoreCase = True
        regex.Global = False
        regex.Pattern = "^1[3|4|5|8]\d{9}$"
        isMobilePhoneNumberValid = regex.Test(mobile)
    End Function
    
    Public Property Get StatusCode()
        StatusCode = status_code
    End Property
    
    Public Property Get StatusMessage()
        StatusMessage = status_message
    End Property
    
    Public Function hasErrorOccured()
        hasErrorOccured = error_occured
    End Function
    
    Public Function login(mobile, password, loginstatus)
        request_devid = mobile
        login = False
        Dim content
        content = post(LOGIN_SUBMIT_DIR, buildLoginParameters(mobile, password, loginstatus))
        If http.status<>200 Then
            Exit Function
        End If
        
        Dim status
        status = parseLoginMessage(content)
        login = status(0)
        
        If Not login Then
            status_code = 1
        Else
            status_code = 0
        End If
        status_message = status(1)
    End Function
    
    Public Function logout()
        logout = False
        
        Dim content
        content = post(LOGOUT_SUBMIT_DIR, "")
        If http.status<>200 Then
            Exit Function
        End If
        
        Dim status
        status = parseLogoutMessage(content)
        logout = status(0)
        
        If Not logout Then
            status_code = 2
        Else
            status_code = 0
        End If
        status_message = status(1)
    End Function
    
    Public Function sendSMS(msg, mobile)
        sendSMS = False
        Dim uid, url
        uid = convertPhoneNumberToUserId(mobile)
        If uid=-1 Then : status_code = 3 : Exit Function
        
        url = buildUrl(SMS_SUBMIT_DIR, buildUserIdParameters(uid))
        Dim content
        content = post(url, buildMessageParameters(msg))
        Dim status
        status = parseSendMessage(content)
        sendSMS = status(0)
        
        If Not sendSMS Then
            status_code = 4
        Else
            status_code = 0
        End If
        status_message = status(1)
    End Function
    
    Public Function sendMessage(msg, mobile)
        sendMessage = False
        Dim uid, url
        uid = convertPhoneNumberToUserId(mobile)
        If uid=-1 Then : status_code = 3 : Exit Function
        
        url = buildUrl(FETION_SUBMIT_DIR, buildUserIdParameters(uid))
        Dim content
        content = post(url, buildMessageParameters(msg))
        Dim status
        status = parseSendMessage(content)
        sendMessage = status(0)
        
        If Not sendMessage Then
            status_code = 4
        Else
            status_code = 0
        End If
        status_message = status(1)
    End Function
    
    Public Function sendMessageToOwn(msg)
        Dim content
        content = post(FETION_MYSELF_SUBMIT_DIR, buildMessageParameters(msg))
        
        Dim status
        status = parseSendMessage(content)
        sendMessageToOwn = status(0)
        
        If Not sendMessageToOwn Then
            status_code = 4
        Else
            status_code = 0
        End If
        status_message = status(1)
    End Function
End Class

Class CommandLineParser
    Private parameters
    Private splitargs
    
    Private Sub Class_Initialize()
        Set parameters = WSH.CreateObject("Scripting.Dictionary")
        Set splitargs = WSH.CreateObject("Scripting.Dictionary")
    End Sub
    
    Private Sub Class_Terminate()
        splitargs.RemoveAll
        Set splitargs = Nothing
        parameters.RemoveAll
        Set parameters = Nothing
    End Sub
    
    Public Sub addSplitter(key, value)
        splitargs.Add key, value
    End Sub
    
    Public Sub parse(args)
        Dim i, j, k, token, c, statusChanged
        Dim statusSkipped, tokens
        Dim keys, items
        
        statusChanged = False
        keys = splitargs.Keys
        items = splitargs.Items
        
        For i = 0 To args.Count-1
            If statusChanged Then
                parameters.Add token, args(i)
                statusChanged = False
            End If
            c = Left(Trim(args(i)), 1)
            If c = "-" Then
                statusChanged = True
                statusSkipped = False
                For j=2 To Len(args(i))-1
                    c = MID(args(i), j, 1)
                    statusSkipped = CBool(c = "-")
                    If Not statusSkipped Then Exit For
                Next
                token = Right(args(i), Len(args(i)) - j+1)
                
                For k = 0 To splitargs.Count - 1
                    tokens = Split(token, items(k))
                
                    If UBound(tokens) > 0 Then
                        If keys(k) = tokens(0) Then
                            parameters.Add tokens(0), _
                                        tokens(1)
                            statusChanged = False
                        End If
                    End If
                Next
            End If
            
            If i = args.Count-1 And statusChanged Then
                parameters.Add token, ""
            End If
        Next
    End Sub
    
    Public Sub dump()
        Dim i, Keys, Items
        Keys = parameters.Keys
        Items = parameters.Items
        
        For i = 0 To parameters.Count - 1
            WSH.Echo Keys(i) & "=" & Items(i)
        Next
    End Sub
    
    Public Function hasArgument(name)
        hasArgument = parameters.Exists(name)
    End Function
    
    Public Function getArgument(name)
        getArgument = ""
        If parameters.Exists(name) Then
            getArgument = parameters(name)
        End If
    End Function
    
    Public Default Property Get Item(name)
        Item = getArgument(name)
    End Property
End Class

'''''''''''''''''''''
' Checks if this script is running under cscript.exe

private function IsCScript()
    if InStrRev(LCase(WScript.FullName), "cscript.exe", -1) <> 0 then
        IsCScript = True
    else 
        IsCScript = False
    end if
end function

Sub DisplayStatusMessage(obj)
    LineOut vbTab & _
         L_Message_Display_OnScreen_Text, Array(obj.StatusMessage, Hex(obj.StatusCode))
End Sub

Sub LineOut(str, args)
    If Not EchoText Then Exit Sub
    If IsNull(args) Then
        WScript.Echo str
    Else
        vbPrintf str, args
    End If
End Sub

Function bool(var)
    If VarType(var) = vbBoolean Then
        bool = var
        Exit Function
    End If
    
    If IsNumeric(var) Then
        var = CInt(var)
        bool = CBool(var <> 0)
    ElseIf TypeName(var) = "String" Then
        var = LCase(var)
        bool = CBool(var = "true" Or var = "enable")
    Else
        bool = False
    End If
End Function

' The following functions vbPrintf and matchPattern based on cmdlib.wsc file
' This file was distributed via Windows XP, but I cannot found it on
' Windows Vista or Windows 7, so I cannot call these two functions
' by CreateObject("Microsoft.CmdLib")

' # Begin Microsoft.CmdLib
' Subroutine which implements normal printf functionality
'********************************************************************
'* Sub:     vbPrintf
'*
'* Purpose: Simulates the Printf function.
'*
'* Input:  [in]  strPhrase      the string with '%1 %2 &3 ' in it
'*         [in]  args           the values to replace '%1 %2 ..etc' with
'*
'* Output:  Displays the string on the screen
'*          (All the '%x' variables in strPhrase is replaced by the 
'*           corresponding elements in the array)
'*
'********************************************************************
Sub vbPrintf(ByVal strPhrase, ByVal args )

    ON ERROR RESUME NEXT
    Err.Clear

    'Changed for localization  

    Dim strMatchPattern         ' the pattern to match - '%[number]'
    Dim intValuesCount          ' to get the count of matching results
    Dim i                       ' used in the loop
    Dim strTemp                 ' to store temporally  the given input string  for formatting

    strTemp   = strPhrase

    ' look out for '%[number]' in the given string
    strMatchPattern = "%\d" '"\%[number]"

    intValuesCount = matchPattern (strMatchPattern, strTemp)

    If intValuesCount <> 0 Then
            ' if present then replace '%1 %2 %3' in the string by
            ' corresponding element in the given array

        If Not IsArray(args) Then
            If IsNull(args) Then
                WScript.Echo(strPhrase)
                Exit Sub
            End If
            args = Array(args)
        End If
               
        If intValuesCount <> UBound(args)+1 Then
            WScript.Echo "L_INVALID_ERRORMESSAGE_ARG_NUMBER_AS_INPUT_ERRORMESSAGE"
            WScript.Quit -10
        End If

        For i = 1 to intValuesCount
            strPhrase = Replace(strPhrase, "%" & Cstr(i), (args(i-1) ), 1, 1, VBBinaryCompare)
        Next

    End If

    WScript.Echo(strPhrase)

End Sub

' Function which checks whether a given value matches a particular pattern
'********************************************************************
'* Function: matchPattern
'*
'* Purpose:  To check if the given pattern is existing in the string
'*
'* Input:
'*  [in]     strMatchPattern   the pattern to look out for
'*  [in]     strPhrase         string in which the pattern needs to be checked
'*
'* Output:   Returns number of occurrences if pattern present, 
'*           Else returns CONST_NO_MATCHES_FOUND
'*
'********************************************************************
Function matchPattern(ByVal strMatchPattern, ByVal strPhrase)

    ON ERROR RESUME NEXT
    Err.Clear

    Dim objRegEx        ' the regular expression object
    Dim Matches         ' the results that match the given pattern
    Dim intResultsCount ' the count of Matches
            
    intResultsCount = 0  ' initialize the count to 0

    'create instance of RegExp object
    Set objRegEx = New RegExp 
    If (NOT IsObject(objRegEx)) Then
        WScript.Echo ("L_ERROR_CHECK_VBSCRIPT_VERSION_ERRORMESSAGE")
    End If
    'find all matches
    objRegEx.Global = True
    'set case insensitive
    objRegEx.IgnoreCase = True
    'set the pattern
    objRegEx.Pattern = strMatchPattern

    Set Matches = objRegEx.Execute(strPhrase)
    intResultsCount = Matches.Count

    'test for match
    If intResultsCount > 0 Then
        matchPattern = intResultsCount
     Else
        matchPattern = -1
    End If

End Function
' # End Microsoft.CmdLib

Sub Usage()
    WScript.Echo L_Help_Help_General01_Text & vbCrLf & _
        L_Help_Help_General02_Text & vbCrLf & _
        L_Help_Help_General03_Text & vbCrLf & _
        L_Help_Help_General04_Text & vbCrLf & _
        L_Help_Help_General05_Text & vbCrLf & _
        L_Help_Help_General06_Text & vbCrLf & _
        L_Help_Help_General07_Text & vbCrLf & _
        L_Help_Help_General08_Text & vbCrLf & _
        L_Help_Help_General09_Text
    WScript.Echo L_Help_Help_General10_Text & vbCrLf & _
        L_Help_Help_General11_Text & vbCrLf & _
        L_Help_Help_General12_Text & vbCrLf & _
        L_Help_Help_General13_Text & vbCrLf & _
        L_Help_Help_General14_Text & vbCrLf & _
        L_Help_Help_General15_Text & vbCrLf & _
        L_Help_Help_General16_Text & vbCrLf & _
        L_Help_Help_General17_Text & vbCrLf & _
        L_Help_Help_General18_Text & vbCrLf & _
        L_Help_Help_General19_Text & vbCrLf & _
        L_Help_Help_General20_Text
End Sub

Function SendFetionMessage( _
        SendPhoneNumber, SendPassword, ReceivePhoneNumber, _
        MessageText, SendType, hasLogin, hasLogout, hasSend, objFetionMessager)
    SendFetionMessage = 0

    If bool(hasSend) And (ReceivePhoneNumber = "" Or _
        ReceivePhoneNumber = SendPhoneNumber) Then
        If Not objFetionMessager.sendMessageToOwn(MessageText) Then
            SendFetionMessage = objFetionMessager.StatusCode
            LineOut L_Message_SendToOwn_Failed_Text, SendPhoneNumber
        Else
            LineOut L_Message_SendToOwn_Succeeded_Text, SendPhoneNumber
        End If
        DisplayStatusMessage objFetionMessager
    ElseIf bool(hasSend) Then
      If objFetionMessager.isMobilePhoneNumberValid(ReceivePhoneNumber) Then
        Select Case SendType
            Case "SMS"
                If Not objFetionMessager.sendSMS(MessageText, ReceivePhoneNumber) Then
                    SendFetionMessage = objFetionMessager.StatusCode
                    LineOut L_Message_SendSMS_Failed_Text, ReceivePhoneNumber
                Else
                    LineOut L_Message_SendSMS_Succeeded_Text, ReceivePhoneNumber
                End If
                DisplayStatusMessage objFetionMessager
            Case Else
                If Not objFetionMessager.sendMessage(MessageText, ReceivePhoneNumber) Then
                    SendFetionMessage = objFetionMessager.StatusCode
                    LineOut L_Message_SendMsg_Failed_Text, ReceivePhoneNumber
                Else
                    LineOut L_Message_SendMsg_Succeeded_Text, ReceivePhoneNumber
                End If
                DisplayStatusMessage objFetionMessager
        End Select
      Else
        SendFetionMessage = 5
        LineOut L_Message_PhoneNumber_Invalid_Text, ReceivePhoneNumber
      End If
    End If
    
End Function

Function ShowOptionBox(msg)
    WScript.StdOut.WriteLine msg
    WScript.StdOut.Write L_Message_Option_Selection_Text
    ShowOptionBox = WScript.StdIn.ReadLine
    If ShowOptionBox = "" Then ShowOptionBox = "Y"
    ShowOptionBox = CBool(UCase(ShowOptionBox) = "Y")
End Function

Function VBMain()
    VBMain = 0
    
    If Not IsCScript() Or WScript.Arguments.Count<1 Then
        Call Usage()
        Exit Function
    End If
    
    Dim objCommandLineParser
    Set objCommandLineParser = New CommandLineParser
        Call objCommandLineParser.addSplitter(L_Argument_HasLogin_Name, L_Argument_Splitter_Token)
        Call objCommandLineParser.addSplitter(L_Argument_HasLogout_Name, L_Argument_Splitter_Token)
        Call objCommandLineParser.addSplitter(L_Argument_TypeName_Name, L_Argument_Splitter_Token)
        Call objCommandLineParser.addSplitter(L_Argument_HasSend_Name, L_Argument_Splitter_Token)
        Call objCommandLineParser.addSplitter(L_Argument_HasEcho_Name, L_Argument_Splitter_Token)
        Call objCommandLineParser.addSplitter(L_Argument_LoginStatus_Name, L_Argument_Splitter_Token)
        
        Call objCommandLineParser.parse(WScript.Arguments)
        ' objCommandLineParser.dump
        
    Dim SendPhoneNumber, ReceivePhoneNumber, SendPassword, SendType
    Dim MessageText, hasLogin, hasLogout, hasSend, LoginStatus
    
    SendPhoneNumber = objCommandLineParser(L_Argument_SendPhoneNumber_Name)
    SendPassword = objCommandLineParser(L_Argument_SendPassword_Name)
    ReceivePhoneNumber = objCommandLineParser(L_Argument_ReceivePhoneNumber_Name)
    MessageText = objCommandLineParser(L_Argument_MessageText_Name)
    hasLogin = objCommandLineParser(L_Argument_HasLogin_Name)
    hasLogout = objCommandLineParser(L_Argument_HasLogout_Name)
    SendType = UCase(objCommandLineParser(L_Argument_TypeName_Name))
    hasSend =  objCommandLineParser(L_Argument_HasSend_Name)
    EchoText = bool(objCommandLineParser(L_Argument_HasEcho_Name))
    LoginStatus = objCommandLineParser(L_Argument_LoginStatus_Name)
    
    If (objCommandLineParser.hasArgument(L_Argument_DisplayHelp_Name)) Then
        Call Usage()
    End If
    
    Select Case LCase(LoginStatus)
        Case "online", "1"
            LoginStatus = "1"
        Case "busy", "2"
            LoginStatus = "2"
        Case "leave", "3"
            LoginStatus = "3"
        Case Else
            LoginStatus = "4"
    End Select
    
    Dim objFetionMessager
    Set objFetionMessager = New FetionMessager
    
    If WScript.Interactive Then
        Do While Not objFetionMessager.isMobilePhoneNumberValid(SendPhoneNumber)
            WScript.StdOut.Write L_Message_Enter_SendNumber_Text
            SendPhoneNumber = WScript.StdIn.ReadLine
            If Not objFetionMessager.isMobilePhoneNumberValid(SendPhoneNumber) Then
                If Not ShowOptionBox(L_Message_PhoneNumber1_Invalid_Text) Then
                    Exit Do
                End If
            End If
        Loop
        
        Do While (SendPassword = "")
            WScript.StdOut.Write L_Message_Enter_SendPassword_Text
            SendPassword = WScript.StdIn.ReadLine
            If SendPassword = "" Then
                If Not ShowOptionBox(L_Message_PhonePassword_Invalid_Text) Then
                    Exit Do
                End If
            End If
        Loop
        
         Do While (MessageText = "")
            WScript.StdOut.WriteLine L_Message_Enter_MessageText_Text
            MessageText = WScript.StdIn.ReadLine
            If MessageText = "" Then
                If Not ShowOptionBox(L_Message_MessageText_Invalid_Text) Then
                    Exit Do
                End If
            End If
         Loop
    End If
    
    If Not objFetionMessager.isMobilePhoneNumberValid(SendPhoneNumber) Then
        VBMain = 5
        LineOut L_Message_PhoneNumber_Invalid_Text, SendPhoneNumber
        Exit Function
    End If
    
    If bool(hasLogin) Then
        If Not objFetionMessager.login(SendPhoneNumber, SendPassword, LoginStatus) Then
            VBMain = objFetionMessager.StatusCode
            LineOut L_Message_Login_Failed_Text, Null
            DisplayStatusMessage objFetionMessager
            Exit Function
        End If
    
        LineOut L_Message_Login_Succeeded_Text, Null
        DisplayStatusMessage objFetionMessager
    End If
    
    Dim i, ReceivePhoneNumbers, LastErrorCode
    ReceivePhoneNumbers = Split(ReceivePhoneNumber, ",")
    LastErrorCode = VBMain
    
    If UBound(ReceivePhoneNumbers) = -1 Then
        VBMain = SendFetionMessage( _
                SendPhoneNumber, SendPassword, ReceivePhoneNumber, _
                MessageText, SendType, hasLogin, hasLogout, hasSend, objFetionMessager)
    Else
        For i = 0 To UBound(ReceivePhoneNumbers)
            VBMain = SendFetionMessage( _
                SendPhoneNumber, SendPassword, ReceivePhoneNumbers(i), _
                MessageText, SendType, hasLogin, hasLogout, hasSend, objFetionMessager)
            If VBMain<>0 Then LastErrorCode = VBMain
        Next
    End If
    
    If bool(hasLogout) Then
        If Not objFetionMessager.logout() Then
            VBMain = objFetionMessager.StatusCode
            LineOut L_Message_Logout_Failed_Text, Null
            DisplayStatusMessage objFetionMessager
            Exit Function
        End If
    
        LineOut L_Message_Logout_Succeeded_Text, Null
        DisplayStatusMessage objFetionMessager
    End If
    
    If VBMain=0 Then VBMain = LastErrorCode
    Set objFetionMessager = Nothing
    Set objCommandLineParser = Nothing
End Function

WScript.Quit(VBMain())
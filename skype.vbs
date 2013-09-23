'''''''''''''''''''''''''''''''''
' File:Skype.vbs
' Date:26/8/2013
' Version: 1.0
' Author: Alban Vidal-Naquet (alban@albanvn.net)
' Sarah plugin for skype
''''''''''''''''''''''''''''''''''
'TODO for v2:
''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Externalize language in text file (all .js and .vbs in one csv file ?)
'PulseEight:Power on the TV + switch TV on Sarah input 
'Script to run Skype.vbs in loop and kill it when NodeJS is down
'Convert Unicode Friend Name from skype to ascii code for xml file
'Minimize skype after connect account done
'message to tell that skype is connected
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Option Explicit
const clsRinging = 4
const cltIncomingP2P = 2
const cltIncomingPSTN = 0
const sarah_tts_url="http://localhost:8080/sarah/parle?phrase="
const sarah_tts_incomingcall="Appelle Skaillpe, de "
const sarah_tts_nocall="Il n'y a pas d'appel disponible"
const sarah_tts_call="Très bien j'appelle "
const sarah_tts_isconnected=" est connécté actuellement, Voulez-vous l'appeler maintenant ?"
const sarah_tts_notconnected=" n'est pas connécté actuellement"
const sarah_tts_notconnected2=" n'est pas connécté actuellement. Je ne peux pas joindre cette personne"
const sarah_tts_notfound="je ne trouve pas dans vos amis ce contact "
const sarah_tts_usernotconnected="La session Skaillpe de l'utilisateur n'est pas ouverte, veuillez l'ouvrir"
const sarah_tts_listconnected="Voici vos amis connectés actuellement."
const sarah_tts_selectfriendok="Votre liste d'amis skaillpe est maintenant à jour"
const sarah_tts_changestatus="Votre statut est maintenant: "
const sarah_tts_getstatus="Votre statut est actuellement: "
const sarah_tts_connected="en ligne"
const sarah_tts_disconnected="hors ligne"
const sarah_tts_busy="occupé"
const sarah_tts_unknownparameter="Je ne reconnais pas le paramètre"
const sarah_tts_notfoundgroup="Je ne trouve pas le groupe"
const sarah_tts_skypeaccount="Connecté au compte skaillpe de "
const sarah_tts_skypenotrunning="Skaillpe n'est pas actuellement lancé"
const sarah_tts_accountnotconnected="Aucun compte connecté actuellement à Skaillpe"
const sarah_tts_thestatusis=" et le statut est: "
const sarah_jsurl="http://127.0.0.1:8080/sarah/skype"

const callin_freqtimer=8000
const callin_repeattimes=3

dim oSkype
dim xmlHttp
dim args
dim StartTime
dim oError
dim oEvent
dim config
dim config_xml
dim cUserStatus_Online 
dim cUserStatus_Offline 
dim cUserStatus_Busy
dim CR
dim g_status
dim g_account
dim g_lastcall

on error resume next

' main()
Set args  = Wscript.Arguments
  if args.count=0 Then
  ' If no argument then start in daemon mode
  StartTime=0
  If CountProcess("wscript.exe", "skype.vbs", "updateparameter") = 1 Then
    init_Skype 
    Do While True  
      WScript.Sleep(60000) 
    Loop
  End If
Else
  'debug purpose
  'send_http_request(sarah_tts_url + args(0))
  init_Skype
  Select case args(0)
    case "updateparameter"
	    Skype_UpdateConfigParameter(args)
    case "answer"
		Skype_Answer()
    case "call"
		Skype_Call(args(1))
    case "callvideo"
		Skype_CallVideo(args(1))
    case "finish"
		Skype_Finish()
    case "videoon"
		Skype_RunVideo()
    case "videooff"
		Skype_StopVideo()
    case "selectfriendsilent"
		Skype_WaitConnexion		
		Skype_SelectFriend args, true
    case "selectfriend"
		Skype_SelectFriend args, false
    case "fullscreen"
		Skype_FullScreen()
    case "cleanwscript"
		Skype_Clean()
    case "isconnected"
		Skype_IsConnected(args(1))
    case "listconnected"
		Skype_ListConnected()
	case "screenon"
	    Skype_ScreenOn()
	case "getstatus"
	    Skype_GetStatus()
	case "connect"
	    Skype_Connect()
	case "disconnect"
	    Skype_Disconnect()
	case "busy"
	    Skype_Busy()
	case "minimize"
		Skype_Minimize()
	case "test"
	    Skype_Test()
	case "undefined"
	    DoNothing()
	case else
	    send_http_request(sarah_tts_url + sarah_tts_unknownparameter + " " + args(0))
  End Select 
End If

Sub Skype_WaitConnexion()
  While Not oSkype.Client.IsRunning
    WScript.Sleep(1000)
  Wend
  oSkype.Attach
  While oSkype.CurrentUserStatus <> cUserStatus_Online
    WScript.Sleep(1000)
  Wend
End Sub

Sub DoNothing()
End Sub

'  Initialize Skype object
Sub init_Skype()
  Set oSkype = WScript.CreateObject("Skype4COM.Skype","oSkype_")
  If Not oSkype.Client.IsRunning Then
     oSkype.Client.Start
  End If
  While Not oSkype.Client.IsRunning
    WScript.Sleep(1000)
  Wend
  oSkype.Attach
  cUserStatus_Online= oSkype.Convert.TextToUserStatus("ONLINE")
  cUserStatus_Offline = oSkype.Convert.TextToUserStatus("OFFLINE")
  cUserStatus_Busy = 4
  CR=Chr(13) + Chr(10)
  g_account=oSkype.CurrentUser.FullName
  g_status=Skype_GetStatusSimple()
  g_lastcall=""
  send_http_request(sarah_jsurl+"?mode=status&account="+g_account+"&status="+g_status+"&lastcall="+g_lastcall)
  ' Check that user is connected
'  If oSkype.CurrentUserStatus <> cUserStatus_Online Then
'	send_http_request(sarah_tts_url + sarah_tts_usernotconnected)
'	WScript.Sleep(2000)
'	WScript.Quit
'  End If
End Sub
 
Sub Skype_Test()
  If Not oSkype.Client.IsRunning Then
     send_http_request(sarah_tts_url+sarah_tts_skypenotrunning)
	 return 
  End If
  If oSkype.CurrentUserStatus <> cUserStatus_Online Then
     send_http_request(sarah_tts_url+sarah_tts_accountnotconnected)
    return
  End If
  If g_status="" Then
    send_http_request(sarah_tts_url+sarah_tts_skypeaccount+g_account)
  Else
    send_http_request(sarah_tts_url+sarah_tts_skypeaccount+g_account+sarah_thestatusis+g_status)
  End If
End Sub

Sub Skype_UpdateConfigParameter(Byval Args)
  FileName=Args(1) + "\\" + "skype.xml"
  content=ReadFile(FileName)
  count=0
  config="	<one-of>"
  For i = 0 To 2
    If (Args(2+i*3)<>"" And Args(2+i*3)<>"undefined") Then
      config = config + CR + "		<item>" + Args(2+i*3) + "<tag>out.action.name=""" + Args(2+i*3) + """;out.action.login=""" + Args(2+i*3+1) + """;out.action.password=""" + Args(2+i*3+2) + """</tag></item>"
	  count=count+1
    End If
  Next
  config = config + CR + "	</one-of>"
  If (count>0) Then
    replaceString=Chr(167) + "3 -->" + CR + config + CR + "	<!-- " + Chr(167) + "3"
    finalstring=ReplacePattern(content, "§3[^§]*§3", replaceString)
    WriteFile FileName, finalstring
  End If
End Sub

' Skype Call Detection
Public Sub oSkype_CallStatus(ByVal pCall , ByVal Status )
       If Status = clsRinging Then
         If Timer() > (StartTime + 10) Then
           StartTime = Timer()
           If pCall.Type = cltIncomingP2P Or pCall.Type = cltIncomingPSTN Then
             For i=1 to callin_repeattimes
	     	   If oSkype.ActiveCalls.Count>0 Then
		         send_http_request(sarah_tts_url + sarah_tts_incomingcall + pCall.PartnerHandle)
  			     send_http_request(sarah_jsurl+"?mode=status&account=&status=&lastcall="+g_lastcall)
     	         Wscript.Sleep(callin_freqtimer)
		       End If
             Next
           End If
         End If
       End If
End Sub

Sub Skype_Connect()
  oSkype.ChangeUserStatus(cUserStatus_Online)
  send_http_request(sarah_tts_url + sarah_tts_changestatus + " " + sarah_tts_connected)
End Sub

Sub Skype_Disconnect()
  oSkype.ChangeUserStatus(cUserStatus_Offline)
  send_http_request(sarah_tts_url + sarah_tts_changestatus + " " + sarah_tts_disconnected)
End Sub

Sub Skype_Busy()
  oSkype.ChangeUserStatus(cUserStatus_Busy)
  send_http_request(sarah_tts_url + sarah_tts_changestatus + " " + sarah_tts_busy)
End Sub

Function Skype_GetStatusSimple()   
  Skype_GetStatusSimple="" 
  If oSkype.CurrentUserStatus = cUserStatus_Online Then
    Skype_GetStatusSimple=sarah_tts_connected
  End If
  If oSkype.CurrentUserStatus = cUserStatus_Offline Then
    Skype_GetStatusSimple=sarah_tts_disconnected
  End If
  If oSkype.CurrentUserStatus = cUserStatus_Busy Then
    Skype_GetStatusSimple=sarah_tts_busy
  End If
End Function

Sub Skype_GetStatus()    
  status=Skype_GetStatusSimple()
  send_http_request(sarah_tts_url + sarah_tts_getstatus + status)
End Sub

Sub Skype_ScreenOn()
	Set oShell = WScript.CreateObject("WScript.Shell")
	oShell.SendKeys("{LEFT}")
End Sub

Sub Skype_ListConnected()
  For Each oFriend In oSkype.Friends
      If oFriend.OnLineStatus = cUserStatus_Online Then
	    str=str+", "+oFriend.FullName
	  End If
  Next
  send_http_request(sarah_tts_url + sarah_tts_listconnected + str)
End Sub

Sub Skype_IsConnected(ByVal Name)
  found=0
  For Each oFriend In oSkype.Friends
    If oFriend.DisplayName=Name Or oFriend.FullName=Name Or oFriend.Handle=Name Then
	  found=1
      If oFriend.OnLineStatus <> cUserStatus_Online Then
	    send_http_request(sarah_tts_url + Name + sarah_tts_notconnected)
	  Else
	    send_http_request(sarah_tts_url + Name + sarah_tts_isconnected)
	  End If
	End If
  Next
  If found=0 Then
    send_http_request(sarah_tts_url + sarah_tts_notfound + Name)
  End If
End Sub

Function CountProcess(ByVal CaptionTitle, ByVal CommandLine, ByVal NotInCommandLine)
  count=0
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select * from Win32_Process where caption='" + CaptionTitle + "'",,48)
  For Each objItem in colItems
	If InStr(objItem.CommandLine, CommandLine) Then
	   If InStr(objItem.CommandLine, NotInCommandLine)=0 Then
	     count=count+1
	   End If
	End If
  Next
  CountProcess=count
End Function

Sub Skype_Clean()
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select * from Win32_Process where caption='wscript.exe'",,48)
  For Each objItem in colItems
	If InStr(objItem.CommandLine, "skype.vbs") Then
		objItem.terminate
	End If
  Next
End Sub

Sub Skype_SendCommand(Byval Command)
  set oCde = oSkype.Command(123,Command,"Retour par défaut",True)
  oSkype.SendCommand(oCde)
'  Wscript.echo "Retour : " & oCde.Reply
End Sub

Sub Skype_Fullscreen()
  strComputer = "."
  Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
  Set colItems = objWMIService.ExecQuery("Select * from Win32_Process where caption='Skype.exe'",,48)
  For Each objItem in colItems
	  Set oShell = WScript.CreateObject("WScript.Shell")
      oSkype.Client.Focus
	  If oShell.AppActivate(objItem.ProcessId) Then
   	        oShell.SendKeys "%{ENTER}"
	  End If
  Next
  Skype_ScreenOn()
End Sub

Function ReadFile(Byval FileName)
    Dim oFso, f
    Set oFso = WScript.CreateObject("Scripting.FileSystemObject")
    Set f = oFso.OpenTextFile(FileName, 1, true, 0)
    ReadFile=f.ReadAll
    f.Close
End Function

Sub WriteFile(Byval FileName, Byval Content)
    Dim oFso, f
	' Save the new xml file
    Set oFso = WScript.CreateObject("Scripting.FileSystemObject")
    Set f = oFso.OpenTextFile(FileName, 2, true, 0)
	f.write(Content)
	f.Close
End Sub

Function ReplacePattern(Byval Content, Byval Pattern, Byval NewString)
	  Set objRegEx = WScript.CreateObject("VBScript.RegExp")
      objRegEx.Global = True   
'      objRegEx.IgnoreCase = True
      objRegEx.Pattern = Pattern
      ReplacePattern = objRegEx.Replace(Content, NewString)
End Function

Sub Skype_SelectFriend(ByVal Args, Byval Silent)
  Dim i, config_xml
  Directory=Args(1)
  GroupName=Args(2)
  found=0
  i=1
  config_xml="	<one-of>"
  For each oGroup In oSkype.CustomGroups
    If oGroup.DisplayName=GroupName Then
	  found=1
      For each oUser In oGroup.Users
	    name=oUser.Handle
	    If oUser.DisplayName <> "" Then
		  name=oUser.DisplayName
		Else
		  If oUser.FullName <> "" Then
		    name=oUser.FullName
	      End If
		End If
		config_xml = config_xml +  CR + "		<item>" + name + "<tag>out.action.name=""" + name + """;</tag></item>"
        i=i+1		
	  Next	
	End If
  Next  
  config_xml=config_xml + CR + "	</one-of>"
  FileName=Directory + "\\" + "skype.xml"
  content=ReadFile(FileName)
  ' Replace in skype.xml the automatic filled friends section with the previous selected items
  replaceString=Chr(167) + "1 -->" + CR + config_xml + CR + "	<!-- " + Chr(167) + "1"
  strNewString=ReplacePattern(content, "§1[^§]*§1", replaceString)
  replaceString=Chr(167) + "2 -->" + CR + config_xml + CR + "	<!-- " + Chr(167) + "2"
  finalstring=ReplacePattern(strNewString, "§2[^§]*§2", replaceString)
  WriteFile FileName, finalstring
  If Silent<>true Then
    If found=0 Then
	  send_http_request(sarah_tts_url + sarah_tts_notfoundgroup + " " + GroupName)
    Else
  	  send_http_request(sarah_tts_url + sarah_tts_selectfriendok)
    End If
  End If
End Sub

Sub Skype_CallVideo(Byval Name)
  Skype_Call(Name)
  WScript.Sleep(2000)
  Skype_RunVideo()
End Sub

Sub Skype_Call(Byval Name)
  found=0
  For Each oFriend In oSkype.Friends
    If oFriend.FullName=Name Or oFriend.DisplayName=Name Or oFriend.Handle=Name Then
      If oFriend.OnLineStatus <> cUserStatus_Online Then
	    send_http_request(sarah_tts_url + Name + sarah_tts_notconnected2)
	  Else
  	    oSkype.PlaceCall(oFriend.Handle)
	  End If
      found=1
      Exit For
    End If
  Next
  If found=0 Then
    send_http_request(sarah_tts_url + sarah_tts_notfound + Name)    
  End If
End Sub

Sub Skype_RunVideo()
  If oSkype.ActiveCalls.Count>0 Then
    oSkype.ActiveCalls.Item(1).StartVideoSend()
    WScript.Sleep(1000)
	Skype_FullScreen()
  Else
    send_http_request(sarah_tts_url + sarah_tts_nocall)
  End If
End Sub

Sub Skype_StopVideo()
  If oSkype.ActiveCalls.Count>0 Then
    oSkype.ActiveCalls.Item(1).StopVideoSend()
	Skype_Minimize()
  Else
    send_http_request(sarah_tts_url + sarah_tts_nocall)
  End If
End Sub

Sub Skype_Minimize()
	oSkype.Client.Minimize()
End Sub

Sub Skype_Finish()
  If oSkype.ActiveCalls.Count>0 Then
    oSkype.ActiveCalls.Item(1).Finish()
	oSkype.Client.Minimize()
  Else
    send_http_request(sarah_tts_url + sarah_tts_nocall)
  End If
End Sub

Sub Skype_Answer()
  If oSkype.ActiveCalls.Count>0 Then
    oSkype.ActiveCalls.Item(1).Answer()
  Else
    send_http_request(sarah_tts_url + sarah_tts_nocall)
  End If
End Sub


' Send http request
Sub send_http_request(ByVal url)
  Set xmlHttp = WScript.CreateObject("MSXML2.ServerXMLHTTP")
  xmlHttp.Open "GET", url, False
  xmlHttp.Send ""
  getHTML = xmlHttp.responseText
  status = xmlHttp.status
  xmlHttp.Abort
End Sub
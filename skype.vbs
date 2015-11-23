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
const sarah_jsurl="http://127.0.0.1:8080/sarah/skype"

const sarah_tts_incomingcall="Appelle Skaillpe, de "
const sarah_tts_nocall="Il n'y a pas d'appel disponible"
const sarah_tts_call="Très bien j'appelle "
const sarah_tts_isconnected=" est connecté actuellement, Voulez-vous l'appeler maintenant ?"
const sarah_tts_notconnected=" n'est pas connecté actuellement"
const sarah_tts_notconnected2=" n'est pas connecté actuellement. Je ne peux pas joindre cette personne"
const sarah_tts_notfound="je ne trouve pas dans vos amis ce contact "
const sarah_tts_usernotconnected="La session Skaillpe de l'utilisateur n'est pas ouverte, veuillez l'ouvrir"
const sarah_tts_listconnected="Voici vos amis connectés actuellement."
const sarah_tts_selectfriendok="Votre liste d'amis skaillpe est maintenant à jour"
const sarah_tts_changestatus="Votre statut est maintenant: "
const sarah_tts_getstatus="Votre statut est actuellement: "
const sarah_tts_connected="en ligne"
const sarah_tts_disconnected="hors ligne"
const sarah_tts_away="absent"
const sarah_tts_invisible="invisible"
const sarah_tts_busy="occupé"
const sarah_tts_unknownparameter="Je ne reconnais pas le paramètre"
const sarah_tts_notfoundgroup="Je ne trouve pas le groupe"
const sarah_tts_skypeaccount="Connecté au compte skaillpe de "
const sarah_tts_skypenotrunning="Skaillpe n'est pas actuellement lancé"
const sarah_tts_accountnotconnected="Aucun compte connecté actuellement à Skaillpe"
const sarah_tts_thestatusis=". Le statut est: "

const CALLIN_FREQTIMER=8000
const CALLIN_REPEATTIMES=3
const g_debug=0

dim debugapp
dim aliveFreq
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
dim cUserStatus_Away 
dim cUserStatus_Invisible 
dim cUserStatus_Busy
dim CR
dim g_status
dim g_account
dim g_directory


If g_debug=0 Then
  on error resume next
End If

Sub Debug_Old(Byval level, Byval str)
  If debugapp And level Then
    Dim oFso, f
    Set oFso = WScript.CreateObject("Scripting.FileSystemObject")
    Set f = oFso.OpenTextFile(g_directory+"\\debuglog.txt", 8, True)
	f.Write str & vbCrLf
	f.Close
  End If
End Sub

Sub Debug(Byval level, Byval str)
  If debugapp And level Then
    send_http_request_debug(sarah_jsurl+"?mode=debug&comment="+str)
  End If
End Sub

' main()
Set args  = Wscript.Arguments
Set fso = WScript.CreateObject("Scripting.FileSystemObject")
Set objFile = fso.GetFile(WScript.ScriptFullName)
g_directory = Left(objFile.Path, Len(objFile.Path)-Len(objFile.Name))
debugapp=0
If args.Count=3 And args(1)="daemon" Then
	' If no argument then start in daemon mode
	StartTime=0
	debugapp=args(0)
	aliveFreq=args(2)
	dim quit
	quit=0
	If CountProcess("wscript.exe", "skype.vbs", "daemon") = 1 Then
	  Debug 16,"Running skype.vbs in foreground"
	  init_Skype
      send_http_request(sarah_jsurl+"?mode=alive")
	  Do While quit=0
		WScript.Sleep(aliveFreq*1000) 
   	    send_http_request(sarah_jsurl+"?mode=alive")
		If CheckNodeJS() = 1 Then
		  quit=1
		End If
	  Loop
	  Debug 16, "Skype.vbs daemon is ending"
	Else
	  Debug 16, "Skype.vbs already running in foreground"
	End If
Else
	If args.Count>1 Then
		init_Skype
		debugapp=args(0)
		Debug 16, "Dispatching msg: mode=" + args(1) + " nbarg=" + CStr(args.Count-1)
		For I = 2 to args.Count
		  Debug 16, "arg "+CStr(I)+": '" + args(I)+"'"
		Next
		Select case args(1)
			case "updateparameter"
				Skype_UpdateConfigParameter(args)
			case "answer"
				Skype_Answer()
			case "call"
				Skype_Call(args(2))
			case "callvideo"
				Skype_CallVideo(args(2))
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
				Skype_IsConnected(args(2))
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
				send_http_request(sarah_tts_url + sarah_tts_unknownparameter + " " + args(1))
		End Select 
	End If
End If
WScript.Quit 0

Function CheckNodeJS()
  CheckNodeJS=0
  If CountProcess("node.exe", "script/", "wsrnode.js") = 0 Then
	CheckNodeJS=1
  End If
End Function

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
  cUserStatus_Away = oSkype.Convert.TextToUserStatus("AWAY")
  cUserStatus_Invisible = oSkype.Convert.TextToUserStatus("INVISIBLE")
  cUserStatus_Busy = 4
  CR=Chr(13) + Chr(10)
  g_account=oSkype.CurrentUser.Handle
  g_status=Skype_GetStatusSimple()
  send_http_request(sarah_jsurl+"?mode=status&account="+g_account+"&status="+g_status)
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
    send_http_request(sarah_tts_url+sarah_tts_skypeaccount+g_account+sarah_tts_thestatusis+g_status)
  End If
End Sub

Sub Skype_UpdateConfigParameter(Byval Args)
  FileName=Args(2) + "\\" + "skype.xml"
  content=ReadFile(FileName)
  count=0
  config="	<one-of>"
  For i = 0 To 2
    If (Args(3+i)<>"" And Args(3+i)<>"undefined") Then
      config = config + CR + "		<item>" & Args(3+i) & "<tag>out.action.account=""" & (i+1) & """;</tag></item>"
	  count=count+1
    End If
  Next
  config = config + CR + "	</one-of>"
  If (count>0) Then
    replaceString=Chr(167) + "3 -->" + CR + config + CR + "	<!-- " + Chr(167) + "3"
    finalstring=ReplacePattern(content, "§3[^§]*§3", replaceString)
    WriteFile FileName, finalstring
  End If
  content=""
End Sub

' Skype Call Detection
Public Sub oSkype_CallStatus(ByVal pCall , ByVal Status )
   If Status = clsRinging Then
	 ref = DateDiff("S", "1/1/1970", Now())
	 If ref > (StartTime + 10) Then
	   StartTime = ref
	   If pCall.Type = cltIncomingP2P Or pCall.Type = cltIncomingPSTN Then
		 For i=1 to CALLIN_REPEATTIMES
		   If oSkype.ActiveCalls.Count>0 Then
			 send_http_request(sarah_jsurl+"?mode=status&lastcall="+pCall.PartnerHandle)
			 send_http_request(sarah_tts_url + sarah_tts_incomingcall + pCall.PartnerHandle)
			 Wscript.Sleep(CALLIN_FREQTIMER)
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
  If oSkype.CurrentUserStatus = cUserStatus_Away Then
    Skype_GetStatusSimple=sarah_tts_away
  End If
  If oSkype.CurrentUserStatus = cUserStatus_Invisible Then
    Skype_GetStatusSimple=sarah_tts_invisible
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
	   If InStr(objItem.CommandLine, NotInCommandLine) Then
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
  Directory=Args(2)
  GroupName=Args(3)
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
  Debug 32, "URL: "+url
  xmlHttp.Open "GET", url, False
  xmlHttp.Send ""
  getHTML = xmlHttp.responseText
  status = xmlHttp.status
  xmlHttp.Abort
End Sub

' Send http request
Sub send_http_request_debug(ByVal url)
  Set xmlHttp = WScript.CreateObject("MSXML2.ServerXMLHTTP")
  xmlHttp.Open "GET", url, False
  xmlHttp.Send ""
  getHTML = xmlHttp.responseText
  status = xmlHttp.status
  xmlHttp.Abort
End Sub
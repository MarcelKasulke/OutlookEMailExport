VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==========================================================================
'Export Outlook EMail to Drive
'--------------------------------------------------------------------------
'Version: 0.3, 2016-07-25
'Autor: Marcel-Alexander Kasulke (David-Software GmbH)
'==========================================================================
Option Explicit
'-------------------------------------------------------------
' Optionen Skript
'-------------------------------------------------------------
' Dieses Outlook Makro ermöglicht die automatische Speicherung von E-Mails in Dateiverzeichnissen mit vorheriger Auswahl oder
' in einem Standardverzeichnis. Der Pfad für das Standardverzeichnis wird durch ein externes Programm erzeugt, die Ticketfee.

'Parameter des Skriptes:
'Email Ausgabeformat:
' MSG = Outlook msg format (incl. attachments, embedded objects etc.)., TXT = plain text
Const EXM_OPT_MAILFORMAT As String = "MSG"
'Datumsformat des Filenames
Const EXM_OPT_FILENAME_DATEFORMAT As String = "yyyy-mm-dd_hh-nn-ss"
'Baue Dateinamen; Platzhalter: <DATE> für Datum, <SENDER> für Sender's Name, <RECEIVER> für Empfänger, <SUBJECT> für Thema/Title
Const EXM_OPT_FILENAME_BUILD As String = "<DATE>_<SUBJECT>"
'Nutze Browse Folder Funktion? Setze diesen Wert auf false wenn der Browsedialog nicht angezeigt werden soll. Damm werden die E-Mails in den Tagesordner gespeichert
'Der Tagesordner erfordert ein externes Programm das die entsprechenden Pfade erzeugt
Const EXM_OPT_USEBROWSER As Boolean = False
'Zielordner (genutzt wenn EXM_OPT_USEBROWSER =false)
Const EXM_OPT_TARGETFOLDER As String = "C:\"
'Zielordner für PSC E-Mails zum speichern
Const EXM_OPT_IM_FOLDER As String = "X:\XXX\01_psc\"
'Maximale Anzahl an E-Mails die ausgewählt und exportiert werden dürfen. Bitte keine sehr große Anzahl eintragen, dies sorgt unter Umständen für Probleme.
'Empfohlen ist ein Wert zwischen 5 und 20.
Const EXM_OPT_MAX_NO As Integer = 20
'Email subject prefixes (wie "RE:", "FW:" etc.) die Entfernt werden sollen vor dem Speichern. Achtung es handelt sich hierbei um einen regulären Ausdruck
'RegEx. Bsp. "\s" bedeutet Leerzeichen " ".
Const EXM_OPT_CLEANSUBJECT_REGEX As String = "RE:\s|Re:\s|AW:\s|FW:\s|WG:\s|SV:\s|Antwort:\s"
' Daily Export Verzeichnis auf Desktop nutzen? Ja/Nein
Const EXM_OPT_DAILY_DIR = False
' Pfad der Exportdatei für Ticketfee, um IMs zu lesen
Const EXM_OPT_PATH_IM_FILE = "C:\temp\im.txt"
'-------------------------------------------------------------

 

'-------------------------------------------------------------
' Übersetzungen der Dialoge
'-------------------------------------------------------------
'-- English
'Const EXM_007 = "Script terminated"
'Const EXM_013 = "Selected Outlook item is not an e-mail"
'Const EXM_014 = "File already exists"
'-- German
Private Const EXM_001 As String = "Die E-Mail wurde erfolgreich abgelegt."
Private Const EXM_002 As String = "Die E-Mail konnte nicht abgelegt werden, Grund:"
Private Const EXM_003 As String = "Ausgewählter Pfad:"
Private Const EXM_004 As String = "E-Mail(s) ausgewählt und erfolgreich abgelegt."
Private Const EXM_005 As String = "<FREE>"
Private Const EXM_006 As String = "<FREE>"
Private Const EXM_007 As String = "Script abgebrochen"
Private Const EXM_008 As String = "Fehler aufgetreten: Sie haben mehr als [LIMIT_SELECTED_ITEMS] E-Mails ausgewählt. Die Aktion wurde beendet."
Private Const EXM_009 As String = "Es wurde keine E-Mail ausgewählt."
Private Const EXM_010 As String = "Es ist ein Fehler aufgetreten: es war keine Email im Fokus, so dass die Ablage nicht erfolgen konnte."
Private Const EXM_011 As String = "Es ist ein Fehler aufgetreten:"
Private Const EXM_012 As String = "Die Aktion wurde beendet."
Private Const EXM_013 As String = "Ausgewähltes Outlook-Dokument ist keine E-Mail"
Private Const EXM_014 As String = "Datei existiert bereits"
Private Const EXM_015 As String = "<FREE>"
Private Const EXM_016 As String = "Bitte wählen Sie den Ordner zum Exportieren:"
Private Const EXM_017 As String = "Fehler beim Exportieren aufgetreten"
Private Const EXM_018 As String = "Export erfolgreich"
Private Const EXM_019 As String = "Bei [NO_OF_FAILURES] E-Mail(s) ist ein Fehler aufgetreten:"
Private Const EXM_020 As String = "[NO_OF_SELECTED_ITEMS] E-Mail(s) wurden ausgewählt und [NO_OF_SUCCESS_ITEMS] E-Mail(s) erfolgreich abgelegt."
'-------------------------------------------------------------
 
 
'-------------------------------------
'Browse-Folder Dialog
'-------------------------------------
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260
Private Type BrowseInfo
 
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
 
 
Public Sub ExportEmailToDrive()
    Const PROCNAME As String = "Export-Email"
 
    On Error GoTo ErrorHandler
 
    Dim myExplorer As Outlook.Explorer
    Dim myfolder As Outlook.MAPIFolder
    Dim myItem As Object
    Dim olSelection As Selection
    Dim strBackupPath As String
    Dim intCountAll As Integer
    Dim intCountFailures As Integer
    Dim strStatusMsg As String
    Dim vSuccess As Variant
    Dim strTemp1 As String
    Dim strTemp2 As String
    Dim strErrorMsg As String
    Dim strDate As String
    Dim dot As String
    Dim enviro As String
    Dim sysuser As String
 
    '-------------------------------------
    'Get target drive
    '-------------------------------------
    dot = "."
    strDate = Now
    'Erzeuge Datumsstring für Ordner die über Tasktrack erstellt wurden/Arbeitsverzeichnis
    sysuser = CStr(Environ("USER"))
    strDate = "C:\Users\sysuser\Desktop\" & (Year(strDate)) & dot & Format(Month(strDate), "00") & dot & Format(Day(strDate), "00")
        
    If (EXM_OPT_USEBROWSER = True) Then
        'Setzte Variable auf Inhalt des Dialogs
        strBackupPath = GetFileDir
        
 
        'Fehler direkt anzeigen
        If Left(strBackupPath, 15) = "ERROR_OCCURRED:" Then
            strErrorMsg = Mid(strBackupPath, 16, 9999)
            Error 5004
        End If
    Else
        'strBackupPath = EXM_OPT_TARGETFOLDER
        'Lese Umgebungsvariable für IM-Tickets, dieser Wert fird über die Ticketfee gesetzt
        'Hier kommt man hin wenn der FileBrowser aus ist, kein Dialog wird eingeblendet
    'strBackupPath = CStr(Environ("IM"))
    strErrorMsg = Mid(strBackupPath, 16, 9999)
 
    End If
    ' Wenn bei aktiviertem FileBrowser Abbrechen/Schließen geklickt wird dann werden die Daten hierhin gespeichert
    'If strBackupPath = " " Then strBackupPath = strDate
    ' Obiger Eintrag gilt nur für den Tasktracker
    'Then GoTo ExitScript
    If Len(Trim(strBackupPath)) = 0 Then
        If EXM_OPT_DAILY_DIR = False Then
                strBackupPath = readIM(EXM_OPT_PATH_IM_FILE)
        End If
        If EXM_OPT_DAILY_DIR = True Then
             strBackupPath = strDate
        End If
    End If
    If (Not Right(strBackupPath, 1) = "\") Then strBackupPath = strBackupPath & "\"
    '-------------------------------------
    ' Prozess der den Fokus auf die ausgewählten E-Mails order Ordner legt
    ' Case 2 funktioniert normalerweise auch bei geöffneten E-Mails aber leider nicht zuverlässig, machmal lassen sich die E-Mails nicht mehr öffnen
    '-------------------------------------
 
    Set myExplorer = Application.ActiveExplorer
    Set myfolder = myExplorer.CurrentFolder
    If myfolder Is Nothing Then Error 5001
    If Not myfolder.DefaultItemType = olMailItem Then GoTo ExitScript
 
    'Stop if more than x emails selected
    If myExplorer.Selection.Count > EXM_OPT_MAX_NO Then Error 5002
 
    'No email selected at all?
    If myExplorer.Selection.Count = 0 Then Error 5003
 
    Set olSelection = myExplorer.Selection
    intCountAll = 0
    intCountFailures = 0
    For Each myItem In olSelection
        intCountAll = intCountAll + 1
        vSuccess = ProcessEmail(myItem, strBackupPath)
        If (Not vSuccess = True) Then
            Select Case intCountFailures
                Case 0: strStatusMsg = vSuccess
                Case 1: strStatusMsg = "1x " & strStatusMsg & Chr(10) & "1x " & vSuccess
                Case Else: strStatusMsg = strStatusMsg & Chr(10) & "1x " & vSuccess
            End Select
            intCountFailures = intCountFailures + 1
        End If
    Next
    If intCountFailures = 0 Then
        strStatusMsg = intCountAll & " " & EXM_004
    End If
 
 
    'Final Message
    If (intCountFailures = 0) Then  'No failure occurred
        MsgBox strStatusMsg & Chr(10) & Chr(10) & EXM_003 & " " & strBackupPath, 64, EXM_018
    ElseIf (intCountAll = 1) Then   'Only one email was selected and a failure occurred
        MsgBox EXM_002 & Chr(10) & vSuccess & Chr(10) & Chr(10) & EXM_003 & " " & strBackupPath, 48, EXM_017
    Else    'More than one email was selected and at least one failure occurred
        strTemp1 = Replace(EXM_020, "[NO_OF_SELECTED_ITEMS]", intCountAll)
        strTemp1 = Replace(strTemp1, "[NO_OF_SUCCESS_ITEMS]", intCountAll - intCountFailures)
        strTemp2 = Replace(EXM_019, "[NO_OF_FAILURES]", intCountFailures)
        MsgBox strTemp1 & Chr(10) & Chr(10) & strTemp2 & Chr(10) & Chr(10) & strStatusMsg _
        & Chr(10) & Chr(10) & EXM_003 & " " & strBackupPath, 48, EXM_017
    End If
 
 
ExitScript:
    Exit Sub
ErrorHandler:
    Select Case Err.Number
    Case 5001:  'Not an email
        MsgBox EXM_010, 64, EXM_007
    Case 5002:
        MsgBox Replace(EXM_008, "[LIMIT_SELECTED_ITEMS]", EXM_OPT_MAX_NO), 64, EXM_007
    Case 5003:
        MsgBox EXM_009, 64, EXM_007
    Case 5004:
        MsgBox EXM_011 & Chr(10) & Chr(10) & strErrorMsg, 48, EXM_007
    Case Else:
        MsgBox EXM_011 & Chr(10) & Chr(10) _
        & Err & " - " & Error$ & Chr(10) & Chr(10) & EXM_012, 48, EXM_007
    End Select
    Resume ExitScript
End Sub
 
'Hier wird das Speichern der E-Mails durchgeführt

Private Function ProcessEmail(myItem As Object, strBackupPath As String) As Variant
    'Saves the e-mail on the drive by using the provided path.
    'Returns TRUE if successful, and FALSE otherwise.
 
    Const PROCNAME As String = "ProcessEmail"
 
    On Error GoTo ErrorHandler
 
    Dim myMailItem As MailItem
    Dim strDate As String
    Dim strSender As String
    Dim strReceiver As String
    Dim strSubject As String
    Dim strFinalFileName As String
    Dim strFullPath As String
    Dim vExtConst As Variant
    Dim vTemp As String
    Dim strErrorMsg As String
 
    If TypeOf myItem Is MailItem Then
         Set myMailItem = myItem
    Else
        Error 1001
    End If
 
    'Set filename
    strDate = Format(myMailItem.ReceivedTime, EXM_OPT_FILENAME_DATEFORMAT)
    strSender = myMailItem.SenderName
    strReceiver = myMailItem.To 'All receiver, semikolon separated string
    If InStr(strReceiver, ";") > 0 Then strReceiver = Left(strReceiver, InStr(strReceiver, ";") - 1)
    strSubject = myMailItem.Subject
    strFinalFileName = EXM_OPT_FILENAME_BUILD
    strFinalFileName = Replace(strFinalFileName, "<DATE>", strDate)
    strFinalFileName = Replace(strFinalFileName, "<SENDER>", strSender)
    strFinalFileName = Replace(strFinalFileName, "<RECEIVER>", strReceiver)
    strFinalFileName = Replace(strFinalFileName, "<SUBJECT>", strSubject)
    strFinalFileName = CleanString(strFinalFileName)
    If Left(strFinalFileName, 15) = "ERROR_OCCURRED:" Then
        strErrorMsg = Mid(strFinalFileName, 16, 9999)
        Error 1003
    End If
    strFinalFileName = IIf(Len(strFinalFileName) > 251, Left(strFinalFileName, 251), strFinalFileName)
    strFullPath = strBackupPath & strFinalFileName
 
    'Save as msg or txt?
    Select Case UCase(EXM_OPT_MAILFORMAT)
        Case "MSG":
            strFullPath = strFullPath & ".msg"
            vExtConst = olMSG
        Case Else:
            strFullPath = strFullPath & ".txt"
            vExtConst = olTXT
    End Select
    'File already exists?
    If CreateObject("Scripting.FileSystemObject").FileExists(strFullPath) = True Then
        Error 1002
    End If
 
    'Save file
    myMailItem.SaveAs strFullPath, vExtConst
 
    'Return true as everything was successful
    ProcessEmail = True
    myMailItem.Delete
 
 
ExitScript:
    Exit Function
ErrorHandler:
    Select Case Err.Number
    Case 1001:  'Not an email
        ProcessEmail = EXM_013
    Case 1002:
        ProcessEmail = EXM_014
    Case 1003:
        ProcessEmail = strErrorMsg
    Case Else:
        ProcessEmail = "Error #" & Err & ": " & Error$ & " (Procedure: " & PROCNAME & ")"
    End Select
    Resume ExitScript
End Function
 
 
Private Function CleanString(strData As String) As String
 
    Const PROCNAME As String = "CleanString"
 
    On Error GoTo ErrorHandler
 
    'Instantiate RegEx
    Dim objRegExp As Object
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True
 
    'Cut out strings we don't like
    objRegExp.Pattern = EXM_OPT_CLEANSUBJECT_REGEX
    strData = objRegExp.Replace(strData, "")
 
    'Replace and cut out invalid strings.
    strData = Replace(strData, Chr(9), "_")
    strData = Replace(strData, Chr(10), "_")
    strData = Replace(strData, Chr(13), "_")
    objRegExp.Pattern = "[/\\*]"
    strData = objRegExp.Replace(strData, "-")
    objRegExp.Pattern = "[""]"
    strData = objRegExp.Replace(strData, "'")
    objRegExp.Pattern = "[:?<>\|]"
    strData = objRegExp.Replace(strData, "")
 
    'Replace multiple chars by 1 char
    objRegExp.Pattern = "\s+"
    strData = objRegExp.Replace(strData, " ")
    objRegExp.Pattern = "_+"
    strData = objRegExp.Replace(strData, "_")
    objRegExp.Pattern = "-+"
    strData = objRegExp.Replace(strData, "-")
    objRegExp.Pattern = "'+"
    strData = objRegExp.Replace(strData, "'")
 
    'Trim
    strData = Trim(strData)
 
    'Return result
    CleanString = strData
 
 
ExitScript:
    Exit Function
ErrorHandler:
    CleanString = "ERROR_OCCURRED:" & "Error #" & Err & ": " & Error$ & " (Procedure: " & PROCNAME & ")"
    Resume ExitScript
End Function
 
Private Function GetFileDir() As String
 
    Const PROCNAME As String = "GetFileDir"
 
    On Error GoTo ErrorHandler
 
    Dim ret As String
    Dim lpIDList As Long
    Dim sPath As String
    Dim udtBI As BrowseInfo
    Dim RdStrings() As String
    Dim nNewFiles As Long
 
    'Show a browse-for-folder form:
    With udtBI
        .lpszTitle = lstrcat(EXM_016, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
 
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList = 0 Then Exit Function
 
    'Get the selected folder.
    sPath = String$(MAX_PATH, 0)
 
 
    SHGetPathFromIDList lpIDList, sPath
    CoTaskMemFree lpIDList
 
    'Strip Nulls
    If (InStr(sPath, Chr$(0)) > 0) Then sPath = Left$(sPath, InStr(sPath, Chr(0)) - 1)
 
    'Return Dir
    GetFileDir = sPath
 
ExitScript:
    Exit Function
ErrorHandler:
    GetFileDir = "ERROR_OCCURRED:" & "Error #" & Err & ": " & Error$ & " (Procedure: " & PROCNAME & ")"
    Resume ExitScript
End Function
Public Function readIM(strDatei As String) As String
 
     Dim objFSO As Object
     Dim objFile As Object
     Dim strText As String
 
        Const ForReading = 1
        Set objFSO = CreateObject("Scripting.FileSystemObject")
        Set objFile = objFSO.GetFile(strDatei)
        If objFile.Size > 0 Then
        Set objFile = objFSO.OpenTextFile(strDatei, ForReading)
        strText = objFile.ReadAll
        objFile.Close
        End If
        readIM = EXM_OPT_IM_FOLDER & strText
End Function




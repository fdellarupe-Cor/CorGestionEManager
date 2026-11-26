<%
Class Syslog
 Private strFile

' Private Prioridad=array(0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23)

  '*********************************************************************
  ' Initialization/Termination
  '*********************************************************************

  'Initialize Class Members
  Private Sub Class_Initialize()
    strFile = ""
  End Sub

  'Terminate and unload all created objects
  Private Sub Class_Terminate()
    Set objDoc = Nothing
  End Sub

  '*********************************************************************
  ' Properties
  '*********************************************************************



  Public Property Get File()
    File = strFile
  End Property

  '*********************************************************************
  ' Functions
  '*********************************************************************

  'Create Blank XML File, set current obj File to newly created file
  Public Function createFile(strPath)
    Dim objFSO, objTextFile
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    Set objTextFile = objFSO.CreateTextFile(strPath, True)
    objTextFile.Close
    Me.File = strPath
    Set objTextFile = Nothing
    Set objFSO = Nothing
  End Function

Public Function SaveMsg(RECURSO,SEVERIDAD,FECHA,XHORA,IP,MSG,XEXTRA)

    MES = Month(FECHA)
    DIA = Day(FECHA)
   ' HORA = Hour(XHORA)
   ' MINUTOS = Minute(XHORA)
   ' SEGUNDOS = Second(XHORA)
    delim= chr(9)
    Select Case MES
        Case 1
            CAD_MES = "Jan"
        Case 2
            CAD_MES = "Feb"
        Case 3
            CAD_MES = "Mar"
        Case 4
            CAD_MES = "Apr"
        Case 5
            CAD_MES = "May"
        Case 6
            CAD_MES = "Jun"
        Case 7
            CAD_MES = "Jul"
        Case 8
            CAD_MES = "Aug"
        Case 9
            CAD_MES = "Sep"
        Case 10
            CAD_MES = "Oct"
        Case 11
            CAD_MES = "Nov"
        Case 12
            CAD_MES = "Dec"
    End Select

    If Len(DIA) <> 2 Then
        DIA = " " & DIA
    End If

    If Len(HORA) <> 2 Then
        HORA = "0" & HORA
    End If

    If Len(MINUTOS) <> 2 Then
        MINUTOS = "0" & MINUTOS
    End If

    If Len(SEGUNDOS) <> 2 Then
        SEGUNDOS = "0" & SEGUNDOS
    End If


    PRIORIDAD=(Recurso * 8) + Severidad

    HOSTNAME = IP
   ' SaveMsg= SLOG_PRI  & CAD_MES & " " & DIA & " " & XHORA & ":" & MINUTOS & ":" & SEGUNDOS & " " & HOSTNAME & " "  & MSG & CHR(13)
     SaveMsg= SLOG_PRI  & CAD_MES & " " & DIA & " " & XHORA &" " & HOSTNAME & " "  & MSG &  Chr(13) & Chr(10)
  end function

end class
%>
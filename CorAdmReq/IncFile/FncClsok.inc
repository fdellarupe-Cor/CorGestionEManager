<%
class fncClsOK

Private m_oOFV 
Private m_oServerVar 
Private m_oRequestVar 
Private m_oResponseVar 
Private m_oSessionVar 

Public KCSRTYP 
Public KCSRCSZ 
Public KCSRLOC 
Public KCSRLOK 
Public KCSRKND 

Public TIX 
Public TIT 
Public IIT 

Public MTFER 
Public MIFER 

Public MTYP 
Public MIYP 

Public MTYK 
Public MIYK 


Public MIYPT 

Public YFMAXTSK 
Public ZXMFMIN 
Public ZXMFMAX 

Dim ATABTSKW(1000, 22)
Dim ATABLAF(200)
Dim WTFCH(4)

Dim ATABLAFER(200)
'Dim ATABPRYW(1500, 110, 35)
Dim ATABPRYW(350, 300, 35)

Dim StrSql
Dim sHTML, sInput, Celda, Fila, Tabla, sImagen





Public Function UserV()

     UserV = OFV.uv

End Function

Public Property Set OFV(ByVal oOFV )
    
     Set m_oOFV = oOFV

End Property

Public Property Get OFV() 
    
     Set OFV = m_oOFV

End Property






Public Function ATABTSK(ind1 , ind2 )

      ATABTSK = ATABTSKW(ind1, ind2)

End Function


Public Function ATABPROY(ind1 , ind2 , ind3 )

      ATABPROY = ATABPRYW(ind1, ind2, ind3)

End Function


Public Function CrearConsultaEx(Xstrsql, ByRef Xconn, XTipoConsulta, Xa)
 
 Dim rs
 Dim XKCSRTYP
 ErrorDesc = ""

 If KCSRCSZ = 0 Then
     KCSRCSZ = 15
     KCSRLOC = 3
     KCSRLOK = 4
     KCSRKND = 3
     KCSRTYP = 1
 End If

   


 If XTipoConsulta = 1 Then
      XKCSRTYP = KCSRTYP
 Else
      XKCSRTYP = 4
 End If

 
 Select Case XKCSRTYP
        Case 1
                Set rs = Server.CreateObject("adodb.recordset")

                rs.CursorLocation = KCSRLOC   ' 3 adUseClient   2 adUseServer
                rs.CursorType = KCSRKND       ' 3 adOpenStatic
                rs.LockType = KCSRLOK         ' 4 adLockBatchOptimistic '— Or 1 use adLockReadOnly
                rs.CacheSize = KCSRCSZ

              ''  RESPONSE.WRITE Xstrsql & "<BR>"
                rs.Open Xstrsql, Xconn
                Set CrearConsultaEx = rs

        Case 2

                Set CrearConsultaEx = Xconn.Execute(Xstrsql)

        Case 4
                Set rs = Server.CreateObject("adodb.command")
                rs.ActiveConnection = Xconn
                rs.CommandText = Xstrsql
                rs.CommandType = 4

                Set CrearConsultaEx = rs.Execute(, a)
End Select
Set rs = Nothing


End Function



Public Function GETPRYCNFG()

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)
        StrSql = "SELECT * FROM PROYCONFIG "
        Set RSX = CrearConsultaEx(StrSql, CN, 1, parametros)
        If Not RSX.EOF Then
           Session("PRCF") = "S"
           Session("DHC") = RSX(2)
           Session("HPP") = RSX(4)
           Session("DPP") = RSX(6)
           Session("SPP") = RSX(8)
           Session("MPP") = RSX(10)
         End If
        Call OFV.cerrarconsulta(RSX)
        
Call OFV.cerrarconn(CN)

End Function

'''''''''''

Public Function GETTABTSK()


    TIX = TIX + 1
    If TIX > 1000 Then TIX = 1

    GETTABTSK = TIX

End Function

'''''''''''


Public Function GETTABPROYND()

    MTYP = MTYP + 1

    GETTABPROYND = MTYP

End Function

Public Function GETTABTSKND()

    MIYPT = MIYPT + 1

    GETTABTSKND = MIYPT

End Function

Public Function ADMFERIADOS(FOPT, FBUS)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Select Case FOPT
       Case "B"

     WFEC = FBUS
     IIT = 1
     If TIT < 9999 Then
        If ATABLAF(IIT) < WFEC Then
           For IIT = IIT To TIT
               If Not (ATABLAF(IIT) < WFEC) Then Exit For
           Next
           If ATABLAF(IIT) > WFEC Then
              ADMFERIADOS = 0
           ElseIf ATABLAF(IIT) = WFEC Then
              ADMFERIADOS = 1
           Else
              ADMFERIADOS = 0
           End If

        ElseIf ATABLAF(IIT) > WFEC Then
           Do While IIT > 0
              If Not (ATABLAF(IIT) > WFEC) Then Exit Do
              IIT = IIT - 1
           Loop
           If IIT = 0 Then
              ADMFERIADOS = 0
              IIT = 1
           ElseIf ATABLAF(IIT) < WFEC Then
              ADMFERIADOS = 0
           ElseIf ATABLAF(IIT) = WFEC Then
              ADMFERIADOS = 1
           Else
              ADMFERIADOS = 0
           End If

        Else

           ADMFERIADOS = 1

        End If


     Else

       ADMFERIADOS = 0

     End If


       Case "C"
       
       Set CN2 = OFV.conectar(OFV.strconn0)

        TIT = 9999
        WFEC = FBUS - 1

        StrSql = "SELECT * FROM FERIADOS WHERE FECHA > " & WFEC & " ORDER BY FECHA "
        Set RSX = CrearConsultaEx(StrSql, CN2, 1, "")
        If Not RSX.EOF Then
           IIT = 0
           Do While Not RSX.EOF
              IIT = IIT + 1
              ATABLAF(IIT) = RSX(1)
           RSX.MoveNext
           Loop
           TIT = IIT
           IIT = 1

         End If
        Call OFV.cerrarconsulta(RSX)
        
        Call OFV.cerrarconn(CN2)


End Select



End Function


Public Function GETDFER(FDES, FHAS)

Dim RSX 
Dim CN2 


   If Not MTFER > 0 Then
       MTFER = 0
       Set CN2 = OFV.conectar(OFV.strconn0)

        StrSql = "SELECT * FROM FERIADOS ORDER BY FECHA"
        Set RSX = CrearConsultaEx(StrSql, CN2, 1, "")
        If Not RSX.EOF Then
            Do While Not RSX.EOF
                MTFER = MTFER + 1
                ATABLAFER(MTFER) = RSX(1)
                If MTFER = 200 Then Exit Do
                RSX.MoveNext
            Loop
        End If
        Call OFV.cerrarconsulta(RSX)
        Call OFV.cerrarconn(CN2)

    End If

        XCANTFER = 0

        WFEC1 = FDES - 1
        WFEC2 = FHAS + 1

        For IFER = 1 To MTFER
            If ATABLAFER(IFER) > WFEC1 And ATABLAFER(IFER) < WFEC2 Then
                XCANTFER = XCANTFER + 1
            End If

            If Not ATABLAFER(IFER) < WFEC2 Then Exit For
        Next

     GETDFER = XCANTFER


End Function


Public Function GETFECHA(FEDES, FDIAS, FTIPO, FCH)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 




IFEC = FEDES

If Session("PRCF") <> "S" Then
   xxxx = GETPRYCNFG()
End If



Select Case FTIPO

       Case 4
              ical = CInt(FDIAS / CLng(Session("DPP")) / 60)


       Case 1
              ical = CInt(FDIAS / CLng(Session("DPP")))

       Case 2
              ical = FDIAS

       Case 3
              ical = FDIAS * CLng(Session("SPP"))

       Case Else
              ical = FDIAS

End Select


Select Case FCH

       Case 0


            ICORR = IFEC + ical

            WTFCH(0) = ical
            WTFCH(1) = ical
            WTFCH(2) = FEDES
            WTFCH(3) = ICORR


       Case 1

            IHAB = 0

            IFECD = CDate(IFEC + CLng(CDate("01-01-1900")))
            papa = DatePart("W", IFECD, 2)


            ii = 0
            ii3 = 0

            If Session("FCLCTP") = "NK" Then
                XD = papa - 1
                ii3 = ical + (Fix((ical + XD) / 5) * 2)

                DF = GETDFER(IFEC, IFEC + ii3)

                If DF > 0 Then
                    ii3 = ical + DF + (Fix((ical + DF + XD) / 5) * 2)

                End If

            Else






            ii3 = GETDHAB(IFEC, ii3)



            Do While ii < ical
               ii3 = ii3 + 1
               ii3 = GETDHAB(IFEC, ii3)
               ii = ii + 1
            Loop



           End If

           IHAB = IFEC + ii3


           WTFCH(0) = ical
           WTFCH(1) = ii3
           WTFCH(2) = FEDES
           WTFCH(3) = IHAB

End Select


End Function

'''''''''''
Public Function GETDHAB(YFEB, YII3)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


   YII3N = YII3
   YDH = 0
   YFEBN = YFEB

   Do While YDH = 0
      YFEBN = YFEB + YII3N
      YFEBND = CDate(YFEBN + CLng(CDate("01-01-1900")))
  
      YFEBNT = DatePart("W", YFEBND, 2)

      If YFEBNT > 5 Or YFEBNT < 1 Then
         YII3N = YII3N + 1
         YDH = 0
      Else
         YFEBNT = ADMFERIADOS("B", YFEBN)
         If YFEBNT = 0 Then
            YDH = 1
         Else
            YII3N = YII3N + 1
            YDH = 0
         End If
      End If
    Loop

   GETDHAB = YII3N


End Function

'''''''''''
Public Function GetfstartvNK(xtsk, xpry)

    Dim RSXP2 

    Dim CN    

    SXFSTART = 0
    SXFSTARTD = 0
       
    Set CN = OFV.conectar(OFV.strconn6)

    StrSql = "Select * From PROYtareasLST where codigo = '" & xpry & "' and TAREA = '" & xtsk & "' AND ESTADO = 0 order by orden "
    Set RSXP2 = CrearConsultaEx(StrSql, CN, 1, "")
         
    If Not RSXP2.EOF Then
        ARRSTR = RSXP2.GetRows()
 
        Call OFV.cerrarconsulta(RSXP2)
        Call OFV.cerrarconn(CN)

        For XIYETP = 0 To UBound(ARRSTR, 2)

            SXFSTARTD = 0

            If ARRSTR(5, XIYETP) <> xtsk Then
               SXFSTARTD = GETFENDNK(ARRSTR(5, XIYETP), ARRSTR(4, XIYETP))

               NDPRY = GETPROYTAB(ARRSTR(4, XIYETP))

               If NDPRY = 999 Then

                  SXFSTARTD = 0

               Else

                  NDTSK = GETTSKTAB(NDPRY, ARRSTR(5, XIYETP))

                  If NDTSK = 999 Or NDTSK > 8000 Then

                     SXFSTARTD = 0

                  Else

                     SXFSTARTD = ATABPRYW(NDPRY, NDTSK, 5)
  
                     TSKSPL = ATABPRYW(NDPRY, NDTSK, 31)

                     If TSKSPL <> "NO" Then
                        SXFSTARTD = GETFENDNK(TSKSPL, ARRSTR(4, XIYETP))

                        NDTSK = GETTSKTAB(NDPRY, TSKSPL)

                        If NDTSK = 999 Or NDTSK > 8000 Then

                        ''   SXFSTARTD = 0

                        Else

                           SXFSTARTD = ATABPRYW(NDPRY, NDTSK, 5)
  
                        End If


                     End If

                  End If


               End If

            End If

            If SXFSTARTD > SXFSTART Then SXFSTART = SXFSTARTD

        Next

    Else

        Call OFV.cerrarconsulta(RSXP2)
        Call OFV.cerrarconn(CN)

    End If

    GetfstartvNK = SXFSTART
       


End Function



Public Function GETFSTARTV(xtsk, xpry)

    Dim RSXP   
    Dim RSXP2  
    Dim RSXP3  
    Dim RSXP15 

    Dim CN 

    If Session("FCLCTP2") = "NK" Then
       
        GETFSTARTV = GetfstartvNK(xtsk, xpry)
      

    Else



        SXFSTART = 0

        Set CN = OFV.conectar(OFV.strconn6)

        StrSql = "Select * From PROYtareasLST where codigo = '" & xpry & "' and TAREA = '" & xtsk & "' AND ESTADO = 0 order by orden "
        Set RSXP2 = CrearConsultaEx(StrSql, CN, 1, "")
        Do While Not RSXP2.EOF
            'busca el protecto en la columna dcodigo
            StrSql = "Select * From PROYecto where codigo = '" & RSXP2(4) & "'  "
            Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
            If Not RSXP.EOF Then
                SXVPLAN = RSXP(17)
            End If
            Call OFV.cerrarconsulta(RSXP)

            If RSXP2(4) = Session("NCliente") Then
                SXVPLAN = Session("VPLAN")
            End If

            SXFSTARTD = 0
            If RSXP2(5) <> xtsk Then
                StrSql = "Select * From PROYTAREAS WHERE TAREA = '" & RSXP2(5) & "' AND VPLAN =  " & SXVPLAN
                Set RSXP3 = CrearConsultaEx(StrSql, CN, 1, "")
                If Not RSXP3.EOF Then

                    If RSXP3(30) <> "NO" Then

                        StrSql = "Select * From PROYTAREAS WHERE TAREA = '" & RSXP3(30) & "' AND VPLAN =  " & SXVPLAN
                        Set RSXP15 = CrearConsultaEx(StrSql, CN, 1, "")
                        If Not RSXP15.EOF Then
                            If CLng(RSXP15(18)) < 2 Then
                                SXFSTARTD = GETFENDV(RSXP15(3), RSXP15(1))
                            Else
                                SXFSTARTD = RSXP15(11)
                            End If

                        Else
                            If CLng(RSXP3(18)) < 2 Then
                                SXFSTARTD = GETFENDV(RSXP3(3), RSXP3(1))
                            Else
                                SXFSTARTD = RSXP3(11)
                            End If

                        End If
                        Call OFV.cerrarconsulta(RSXP15)

                    Else



                        If CLng(RSXP3(18)) < 2 Then
                            SXFSTARTD = GETFENDV(RSXP3(3), RSXP3(1))
                        Else
                            SXFSTARTD = RSXP3(11)
                        End If


                    End If


                End If
     
                Call OFV.cerrarconsulta(RSXP3)

                If SXFSTARTD > SXFSTART Then SXFSTART = SXFSTARTD
            End If


            RSXP2.MoveNext
        Loop
  
        Call OFV.cerrarconsulta(RSXP2)
  
        Call OFV.cerrarconn(CN)


        GETFSTARTV = SXFSTART
    End If



End Function

Public Function GETFENDV(xtsk, xpry)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 

If Session("FCLCTP") = "NK" Then

    GETFENDV = GETFENDNK(xtsk, xpry)

Else





Set CN = OFV.conectar(OFV.strconn6)

    YXMFMIN = 0
    YXMFMAX = 0
    YXFEND = 0
    YXSALE = "N"
    YXFSTART = CLng(Date) - CLng(CDate("01-01-1900"))
    YXFCURR = YXFSTART
    YXFHOY = YXFSTART
    YXFREAL = YXFSTART

    StrSql = "Select * From PROYecto where codigo = '" & xpry & "'  "
    Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
    If Not RSXP.EOF Then
       If RSXP(4) > 0 Then
          YXFSTART = RSXP(4)
       Else
          YXFSTART = RSXP(3)
       End If
       YXFSTS = RSXP(11)
       YXFFCH = RSXP(18)
       YXVPLAN = RSXP(17)
       If CLng(YXFFCH) > 0 Then XXX = ADMFERIADOS("C", 0)
 
    End If
    Call OFV.cerrarconsulta(RSXP)

    If xpry = Session("NCliente") Then
       YXVPLAN = Session("VPLAN")
    End If


    YXFCURR = YXFSTART
    YXFREAL = YXFSTART

  StrSql = "Select * From PROYetapa where codigo = '" & xpry & "' AND VPLAN = " & YXVPLAN & "  order by orden "
  Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
  If Not RSXP.EOF Then
     Do While Not RSXP.EOF


  StrSql = "Select * From PROYtareas where codigo = '" & xpry & "' and etapa = '" & RSXP(2) & "' AND VPLAN = " & YXVPLAN & " order by orden "

  Set RSXP1 = CrearConsultaEx(StrSql, CN, 1, "")
  If Not RSXP1.EOF Then
     Do While Not RSXP1.EOF



    XIIX = GETTSKCLC(RSXP1, YXFHOY, YXFREAL, YXFFCH)




    YXFREAL = ATABTSKW(XIIX, 1)
    YTSKFS = ATABTSKW(XIIX, 2)
    YTSKFE = ATABTSKW(XIIX, 3)
    YTSKFSR = ATABTSKW(XIIX, 4)
    YTSKFER = ATABTSKW(XIIX, 5)
    YTSKST = ATABTSKW(XIIX, 6)
    YFSPLAN = ATABTSKW(XIIX, 7)
    YFEPLAN = ATABTSKW(XIIX, 8)
    YFSREAL = ATABTSKW(XIIX, 9)
    YFEREAL = ATABTSKW(XIIX, 10)
    YXFFDU = ATABTSKW(XIIX, 11)
    YXFFTI = ATABTSKW(XIIX, 12)
    YPADV = ATABTSKW(XIIX, 13)
    YXAGRD = ATABTSKW(XIIX, 14)
    YXAGRR = ATABTSKW(XIIX, 15)
    YXMODA = ATABTSKW(XIIX, 16)
    YXTDUR = ATABTSKW(XIIX, 17)
    YXUDUR = ATABTSKW(XIIX, 18)
    YXTDUP = ATABTSKW(XIIX, 19)
    YXUDUP = ATABTSKW(XIIX, 20)

    If YXMFMIN = 0 Or YXMFMIN > YTSKFSR Then
       YXMFMIN = YTSKFSR
    End If

   YXFCURR = YXFREAL
   YXFEND = YXFCURR


    If YXMFMAX < YXFREAL Then YXMFMAX = YXFREAL

    If RSXP1(3) = xtsk Then
       YXFEND = YXFCURR
       YXSALE = "S"
       Exit Do
    End If

    RSXP1.MoveNext
    Loop
 
  End If

    If YXSALE = "S" Then
       Exit Do
    End If

    RSXP.MoveNext
    Loop
 
  End If

  If YFMAXTSK = xtsk Then
     ZXMFMIN = YXMFMIN
     ZXMFMAX = YXMFMAX
  End If

  GETFENDV = YXFEND
  
    Call OFV.cerrarconsulta(RSXP1)
    Call OFV.cerrarconsulta(RSXP)
    
    Call OFV.cerrarconn(CN)

 End If

End Function

'''''''''''''''''''''''''''

Public Function GETFENDNK(xtsk, xpry)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 

Dim TYP 
Dim IYP 

Dim IYPT 



    YXMFMIN = 0
    YXMFMAX = 0
    YXFEND = 0
    YXSALE = "N"
    YXFSTART = CLng(Date) - CLng(CDate("01-01-1900"))
    YXFCURR = YXFSTART
    YXFHOY = YXFSTART
    YXFREAL = YXFSTART

    MVETP = 0
    MVTSK = 0
    IYPT = 0

    NDPRY = GETPROYTAB(xpry)

    If NDPRY = 999 Then





   TYP = GETTABPROYND()
   IYP = TYP

   ATABPRYW(IYP, 0, 0) = xpry
   ATABPRYW(IYP, 0, 1) = 0  ' CANT DE TAREAS...

   Set CN = OFV.conectar(OFV.strconn6)


    StrSql = "Select * From PROYecto where codigo = '" & xpry & "'  "
    Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
    If Not RSXP.EOF Then
        If RSXP(4) > 0 Then
            YXFSTART = RSXP(4)
        Else
            YXFSTART = RSXP(3)
        End If
        YXFSTS = RSXP(11)
        YXFFCH = RSXP(18)
        YXVPLAN = RSXP(17)
        YXNOM = RSXP(2)
 
    End If
    Call OFV.cerrarconsulta(RSXP)
    Call OFV.cerrarconn(CN)

    If xpry = Session("NCliente") Then
        YXVPLAN = Session("VPLAN")
    End If

    ATABPRYW(IYP, 0, 2) = YXVPLAN
    ATABPRYW(IYP, 0, 3) = YXFSTS
    ATABPRYW(IYP, 0, 4) = YXFFCH
    ATABPRYW(IYP, 0, 5) = YXFSTART
    ATABPRYW(IYP, 0, 6) = 0
    ATABPRYW(IYP, 0, 10) = YXNOM

  Else


    IYP = NDPRY

    YXVPLAN = ATABPRYW(IYP, 0, 2)
    YXFSTS = ATABPRYW(IYP, 0, 3)
    YXFFCH = ATABPRYW(IYP, 0, 4)
    YXFSTART = ATABPRYW(IYP, 0, 5)


  End If

  MACCPRY = "N"
  MACCTSK = "N"

  YXFCURR = YXFSTART
  YXFREAL = YXFSTART

  If NDPRY <> 999 Then

      MACCPRY = "S"


      NDTSK = GETTSKTAB(NDPRY, xtsk)

      If NDTSK = 999 Then
          MACCTSK = "N"

      ElseIf NDTSK > 8000 Then

              MACCTSK = "P"
              NDTSK = ATABPRYW(IYP, 0, 1)

      Else

              MACCTSK = "S"

      End If

      If MACCTSK = "S" Or MACCTSK = "P" Then

          IYPT = NDTSK

          YXFREAL = ATABPRYW(IYP, IYPT, 1)
          YTSKFSR = ATABPRYW(IYP, IYPT, 4)

          YXMFMIN = ATABPRYW(IYP, IYPT, 23)
          YXMFMAX = ATABPRYW(IYP, IYPT, 24)

          MVETPA = ATABPRYW(IYP, IYPT, 25)
          MVTSKA = ATABPRYW(IYP, IYPT, 26)

          YXFCURR = YXFREAL
          YXFEND = YXFCURR


          WIIX = GETTABTSK()
        
        
          ATABTSKW(WIIX, 0) = ATABPRYW(IYP, IYPT, 0)
          ATABTSKW(WIIX, 1) = ATABPRYW(IYP, IYPT, 1)
          ATABTSKW(WIIX, 2) = ATABPRYW(IYP, IYPT, 2)
          ATABTSKW(WIIX, 3) = ATABPRYW(IYP, IYPT, 3)
          ATABTSKW(WIIX, 4) = ATABPRYW(IYP, IYPT, 4)
          ATABTSKW(WIIX, 5) = ATABPRYW(IYP, IYPT, 5)
          ATABTSKW(WIIX, 6) = ATABPRYW(IYP, IYPT, 6)
          ATABTSKW(WIIX, 7) = ATABPRYW(IYP, IYPT, 7)
          ATABTSKW(WIIX, 8) = ATABPRYW(IYP, IYPT, 8)
          ATABTSKW(WIIX, 9) = ATABPRYW(IYP, IYPT, 9)
          ATABTSKW(WIIX, 10) = ATABPRYW(IYP, IYPT, 10)
          ATABTSKW(WIIX, 11) = ATABPRYW(IYP, IYPT, 11)
          ATABTSKW(WIIX, 12) = ATABPRYW(IYP, IYPT, 12)
          ATABTSKW(WIIX, 13) = ATABPRYW(IYP, IYPT, 13)
          ATABTSKW(WIIX, 14) = ATABPRYW(IYP, IYPT, 14)
          ATABTSKW(WIIX, 15) = ATABPRYW(IYP, IYPT, 15)
          ATABTSKW(WIIX, 16) = ATABPRYW(IYP, IYPT, 16)
          ATABTSKW(WIIX, 17) = ATABPRYW(IYP, IYPT, 17)
          ATABTSKW(WIIX, 18) = ATABPRYW(IYP, IYPT, 18)
          ATABTSKW(WIIX, 19) = ATABPRYW(IYP, IYPT, 19)
          ATABTSKW(WIIX, 20) = ATABPRYW(IYP, IYPT, 20)
          ATABTSKW(WIIX, 21) = ATABPRYW(IYP, IYPT, 21)
          ATABTSKW(WIIX, 22) = ATABPRYW(IYP, IYPT, 22)
        

          IYPT = ATABPRYW(IYP, 0, 1)


       End If





  Else

     NDTSK = 999

     MACCPRY = "N"
     MACCTSK = "N"

  End If


  If MACCPRY = "N" Or (MACCPRY = "S" And (MACCTSK = "N" Or MACCTSK = "P")) Then



  Set CN = OFV.conectar(OFV.strconn6)

  StrSql = "Select * From PROYetapa where codigo = '" & xpry & "' AND VPLAN = " & YXVPLAN & "  order by orden "
  Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
  If Not RSXP.EOF Then

     If MACCTSK = "P" And MVETPA > 0 Then
          RSXP.Move MVETPA
          MVETP = MVETPA
      End If

   End If

   If Not RSXP.EOF Then

      ARRETP = RSXP.GetRows()

     Call OFV.cerrarconsulta(RSXP)
     Call OFV.cerrarconn(CN)

      For XIYETP = 0 To UBound(ARRETP, 2)

     MVTSK = 0

   Set CN = OFV.conectar(OFV.strconn6)

  StrSql = "Select * From PROYtareas where codigo = '" & xpry & "' and etapa = '" & ARRETP(2, XIYETP) & "' AND VPLAN = " & YXVPLAN & " order by orden "

  Set RSXP1 = CrearConsultaEx(StrSql, CN, 1, "")
  If Not RSXP1.EOF Then

     If MACCTSK = "P" And MVTSKA > 0 Then
          RSXP1.Move MVTSKA
          MVTSK = MVTSKA
           MVTSKA = 0
      End If

   End If

   If Not RSXP1.EOF Then

      ARRTSK = RSXP1.GetRows()

     Call OFV.cerrarconsulta(RSXP1)
     Call OFV.cerrarconn(CN)

      For XIYTSK = 0 To UBound(ARRTSK, 2)

     MVTSK = MVTSK + 1


    XIIX = GETTSKCLCNK(ARRTSK, XIYTSK, YXFHOY, YXFREAL, YXFFCH)


 IYPT = IYPT + 1

 ATABPRYW(IYP, 0, 1) = IYPT ' CANT DE TAREAS...

 ATABPRYW(IYP, IYPT, 0) = ARRTSK(3, XIYTSK)

  ATABPRYW(IYP, IYPT, 30) = ARRTSK(5, XIYTSK) ' RSXP1(5)
  ATABPRYW(IYP, IYPT, 29) = MACCTSK
  ATABPRYW(IYP, IYPT, 28) = GETTABTSKND()

  ATABPRYW(IYP, IYPT, 31) = ARRTSK(30, XIYTSK) ' RSXP1(30)

 ATABPRYW(IYP, IYPT, 1) = ATABTSKW(XIIX, 1)
 ATABPRYW(IYP, IYPT, 2) = ATABTSKW(XIIX, 2)
 ATABPRYW(IYP, IYPT, 3) = ATABTSKW(XIIX, 3)
 ATABPRYW(IYP, IYPT, 4) = ATABTSKW(XIIX, 4)
 ATABPRYW(IYP, IYPT, 5) = ATABTSKW(XIIX, 5)
 ATABPRYW(IYP, IYPT, 6) = ATABTSKW(XIIX, 6)
 ATABPRYW(IYP, IYPT, 7) = ATABTSKW(XIIX, 7)
 ATABPRYW(IYP, IYPT, 8) = ATABTSKW(XIIX, 8)
 ATABPRYW(IYP, IYPT, 9) = ATABTSKW(XIIX, 9)
 ATABPRYW(IYP, IYPT, 10) = ATABTSKW(XIIX, 10)
 ATABPRYW(IYP, IYPT, 11) = ATABTSKW(XIIX, 11)
 ATABPRYW(IYP, IYPT, 12) = ATABTSKW(XIIX, 12)
 ATABPRYW(IYP, IYPT, 13) = ATABTSKW(XIIX, 13)
 ATABPRYW(IYP, IYPT, 14) = ATABTSKW(XIIX, 14)
 ATABPRYW(IYP, IYPT, 15) = ATABTSKW(XIIX, 15)
 ATABPRYW(IYP, IYPT, 16) = ATABTSKW(XIIX, 16)
 ATABPRYW(IYP, IYPT, 17) = ATABTSKW(XIIX, 17)
 ATABPRYW(IYP, IYPT, 18) = ATABTSKW(XIIX, 18)
 ATABPRYW(IYP, IYPT, 19) = ATABTSKW(XIIX, 19)
 ATABPRYW(IYP, IYPT, 20) = ATABTSKW(XIIX, 20)
 ATABPRYW(IYP, IYPT, 21) = ATABTSKW(XIIX, 21)
 ATABPRYW(IYP, IYPT, 22) = ATABTSKW(XIIX, 22)

 YXFREAL = ATABPRYW(IYP, IYPT, 1)
 YTSKFSR = ATABPRYW(IYP, IYPT, 4)


    If YXMFMIN = 0 Or YXMFMIN > YTSKFSR Then
        YXMFMIN = YTSKFSR
    End If

   YXFCURR = YXFREAL
   YXFEND = YXFCURR


    If YXMFMAX < YXFREAL Then YXMFMAX = YXFREAL

 ATABPRYW(IYP, IYPT, 23) = YXMFMIN
 ATABPRYW(IYP, IYPT, 24) = YXMFMAX
 ATABPRYW(IYP, IYPT, 25) = MVETP
 ATABPRYW(IYP, IYPT, 26) = MVTSK


    If ARRTSK(3, XIYTSK) = xtsk Then
        YXFEND = YXFCURR
        YXSALE = "S"
        Exit For
    End If


   Next


   Else

     Call OFV.cerrarconsulta(RSXP1)
     Call OFV.cerrarconn(CN)

 
  End If

    If YXSALE = "S" Then
        Exit For
    End If

     MVETP = MVETP + 1

      Next


   Else

     Call OFV.cerrarconsulta(RSXP)
     Call OFV.cerrarconn(CN)
 
  End If



   End If



  If YFMAXTSK = xtsk Then
      ZXMFMIN = YXMFMIN
      ZXMFMAX = YXMFMAX
  End If

  GETFENDNK = YXFEND
  


End Function

'''''''''''''''''''''''''
Public Function GETPROYTAB(xpry)

    NDPRY = 999


    For NDIYP = 1 To MTYP
        If ATABPRYW(NDIYP, 0, 0) = xpry Then
            NDPRY = NDIYP
            Exit For
        End If
    Next

    GETPROYTAB = NDPRY

End Function

'''''''''''''''''''''''''
Public Function GETTSKTAB(NDPRY, xtsk)

    NDTSK = 999

    If ATABPRYW(NDPRY, 0, 1) > 0 Then

        NDTSK = ATABPRYW(NDPRY, 0, 1)

        For NDIYPT = 1 To NDTSK
            If ATABPRYW(NDPRY, NDIYPT, 0) = xtsk Then
                NDTSK = NDIYPT
                Exit For
            End If
        Next

        If NDIYPT > NDTSK Then
            NDTSK = NDTSK + 8000
        End If

    End If

    GETTSKTAB = NDTSK

End Function

'''''''''''''''''''''''''''
Public Function GENPROP()

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)




        StrSql = "INSERT INTO PROYPRTY (CODIGO,CANT,TIPO,ESTR,CONF) VALUES ("
        StrSql = StrSql & "'" & Session("NCliente") & "','0','0','1','0')"
        Call CrearConsultaEx(StrSql, CN, 1, "")


Call OFV.cerrarconn(CN)


End Function

Public Function GENINITPRY()

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)


        StrSql = "Select * From PROYPRTY where codigo = '" & Session("NCliente") & "'  "
        Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
        If Not RSXP.EOF Then

           If RSXP(3) = 0 Then

              YXETP = GENPRYETP(Session("NCliente"), Session("VPLAN"), "10", "ETAPA UNICA", "0", "0")
              YXTSK = GENPRYTSK(Session("NCliente"), Session("VPLAN"), "10", YXETP, "INICIO", "0", "0", "1", "1", "1", "1")


           Else

              YXETP = GENPRYETP(Session("NCliente"), Session("VPLAN"), "10", "INICIAL", "0", "0")
              YXTSK = GENPRYTSK(Session("NCliente"), Session("VPLAN"), "10", YXETP, "INICIO", "0", "0", "1", "1", "1", "1")

              For YII = 2 To RSXP(2) + 1

                  YXETP = GENPRYETP(Session("NCliente"), Session("VPLAN"), (YII * 10), YII, "0", "0")

              Next

              YXETP = GENPRYETP(Session("NCliente"), Session("VPLAN"), (YII * 10), "FINAL", "0", "0")
              YXTSK = GENPRYTSK(Session("NCliente"), Session("VPLAN"), "10", YXETP, "FINALIZACION", "0", "0", "1", "1", "1", "1")

           End If

        End If
       Call OFV.cerrarconsulta(RSXP)
       
       Call OFV.cerrarconn(CN)

End Function

Public Function GENPRYETP(YPRY, YPLAN, YXORDEN, YXNOMBRE, YXSTART, YXEND)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

         YXETAPA = GTETP() & YXORDEN

        StrSql = "INSERT INTO PROYETAPA (CODIGO,ETAPA,ORDEN,DESCRIPCION,FECHASPLAN,FECHASREAL,FECHAEPLAN,FECHAEREAL,VPLAN) VALUES ("
        StrSql = StrSql & "'" & YPRY & "','" & YXETAPA & "','" & YXORDEN & "','" & YXNOMBRE & "'," & YXSTART & ",0," & YXEND & ",0," & YPLAN & ")"
        Call CrearConsultaEx(StrSql, CN, 1, "")

        GENPRYETP = YXETAPA
        
Call OFV.cerrarconn(CN)

End Function

Public Function GENPRYTSK(YPRY, YPLAN, YXORDEN, YXETAPA, YXNOMBRE, YXSTART, YXEND, YXTDUR, YXUDUR, YXTFEC, YTIPO)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

         YXTSK = GTTSK() & YXORDEN

         YXUSR = "NO"

         If YXORDEN = "0" Then YXUSR = Session("USUARIO")

        StrSql = "INSERT INTO PROYTAREAS (CODIGO,ETAPA,TAREA,ORDEN,DESCRIPCION,MODALSEC,MODALINI,"
        StrSql = StrSql & "FECHASPLAN,FECHASREAL,FECHAEPLAN,FECHAEREAL,SECTRESP,USRRESP,TIPO,"
        StrSql = StrSql & "TDUR,UDUR,TIPOFEC,ESTADO,VPLAN,TASIG,CASIG,"
        StrSql = StrSql & "MASIG,PROYTSKTIPO,HORAIRESOL,HORAFRESOL,TSKSPL,TSKCUENTA,MODCREATE)"
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & "'" & YPRY & "','" & YXETAPA & "','" & YXTSK & "','" & YXORDEN & "','" & YXNOMBRE & "',1,0,"
        StrSql = StrSql & YXSTART & ",0," & YXEND & ",0,0,'" & YXUSR & "','NO',"
        StrSql = StrSql & "'" & YXTDUR & "','" & YXUDUR & "','0','0'," & YPLAN & ",'0','0',"
        StrSql = StrSql & "'0','" & YTIPO & "','000000','000000','NO','" & SESSION("PRYCTA") & "','0')"
        Call CrearConsultaEx(StrSql, CN, 1, "")

        GENPRYTSK = YXTSK
        
Call OFV.cerrarconn(CN)

End Function






''''''''''''''''''''''''''
Public Function GENPRYTSKTMP(YPRY, YPLAN, YXORDEN, YXETAPA, YXNOMBRE, YXSTART, YXEND, YXTDUR, YXUDUR, YXTFEC, YTIPO,YCUENTA)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

         YXTSK = GTTSK() & YXORDEN

         YXUSR = Session("USUARIO")

        StrSql = "INSERT INTO PROYTAREASTMP (CODIGO,ETAPA,TAREA,ORDEN,DESCRIPCION,MODALSEC,MODALINI,"
        StrSql = StrSql & "FECHASPLAN,FECHASREAL,FECHAEPLAN,FECHAEREAL,SECTRESP,USRRESP,TIPO,"
        StrSql = StrSql & "TDUR,UDUR,TIPOFEC,ESTADO,VPLAN,TASIG,CASIG,"
        StrSql = StrSql & "MASIG,PROYTSKTIPO,HORAIRESOL,HORAFRESOL,TSKSPL,TSKCUENTA,MODCREATE)"
        StrSql = StrSql & " VALUES ("
        StrSql = StrSql & "'" & YPRY & "','" & YXETAPA & "','" & YXTSK & "','" & YXORDEN & "','" & YXNOMBRE & "',0,0,"
        StrSql = StrSql & "" & YXSTART & "," & YXSTART & "," & YXEND & "," & YXEND & ",0,'" & YXUSR & "','NO',"
        StrSql = StrSql & "'" & YXTDUR & "','" & YXUDUR & "','0','0'," & YPLAN & ",'0','0',"
        StrSql = StrSql & "'0','" & YTIPO & "','000000','000000','NO','" & YCUENTA & "','1')"
        Call CrearConsultaEx(StrSql, CN, 1, "")

        GENPRYTSKTMP = YXTSK
        
Call OFV.cerrarconn(CN)

End Function

''''''''''''''''''''''''''''''''''''''
' IMPUTACION A CUENTAS DE PRESUPUESTO
'''''''''''''''''''''''''''''''''''''''
Public Function GENTSKCUENTA(YXSTART,YXTDUR, YXUDUR, YCUENTA)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim XCN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set XCN = OFV.conectar(OFV.strconn6)

  YXCUENTADET = 0


  StrSql = "Select ID From PROYcuentasDET WHERE CODCUENTA = 0" & YCUENTA
  strsql = strsql & " AND NOT (FINICIO > " & YXSTART & ") AND NOT (FFINAL < " & YXSTART & ") "
  strsql = strsql & " ORDER BY ACTIVA "
  set rsX = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if not rsX.eof then

     YXCUENTADET = RSX(0)

  ELSE


     XCODORDEN = 0


     StrSql = "Select MAX(ORDEN) From PROYcuentasDET WHERE CODCUENTA = 0" & YCUENTA
     set rsXP = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
     if not rsXP.eof then XCODORDEN = RSXP(0)
     call ofv.cerrarconsulta(rsXP)

     XCODORDEN = XCODORDEN + 10
     YXTSK     = GTTSK()
     YXFDHOY   = CLNG(DATE()) - clng(cdate("01-01-1900"))


     strsql = "insert into PROYCUENTASDET (CODCUENTA,DESCRIPCION,ORDEN,FINICIO,FFINAL,ACTIVA,TPLAN,TCONS) "
     strsql = strsql & "values ('" & YCUENTA & "','" & YXTSK & "','" & XCODORDEN & "','" & YXSTART & "','99999','1','1','0')"
    
     call ofv.crearconsultaEx(StrSql,Xcn,1,parametros)


     StrSql = "Select ID From PROYcuentasDET WHERE DESCRIPCION = '" & YXTSK & "'"
     set rsXP = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
     if not rsXP.eof then
        YXCUENTADET = RSXP(0)
     END IF
     call ofv.cerrarconsulta(rsXP)


  END IF
  call ofv.cerrarconsulta(rsX)

     YXCUENTASUM = 0

''''''''YXTDUR, YXUDUR


     SELECT CASE CSTR(YXUDUR & "")

            CASE "2"

                 YXCUENTASUM = YXTDUR * CLng(Session("DPP"))


            CASE "3"

                 YXCUENTASUM = YXTDUR * CLng(Session("DPP")) * CLng(Session("SPP"))


            CASE "4"

                 YXCUENTASUM = REPLACE(CSTR(YXTDUR / 60),",",".")



            CASE ELSE

                 YXCUENTASUM = YXTDUR


      END SELECT


     strsql = "UPDATE PROYCUENTASDET SET TCONS = (TCONS + " & YXCUENTASUM & ") WHERE ID = 0" & YXCUENTADET
     call ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
        
     Call OFV.cerrarconn(XCN)


     GENTSKCUENTA = YXCUENTADET


End Function

''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''
' TRANSFERENCIA DE IMPUTACION A CUENTAS DE PRESUPUESTO
'''''''''''''''''''''''''''''''''''''''
Public Function GENTSKCUENTATRF(YCUENTA)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim XCN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set XCN = OFV.conectar(OFV.strconn6)


  StrSql = "Select * From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
  set rsX = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
  if not rsX.eof then

'''''''''  TRANSFERENCIA DE IMPUTACIONES DE LA TAREA '''''''''''''''''''''''''''''''''''''''''''''


     YCUENTAOLD = RSX(32)


'''''''''' LOOP DE IMPUTACIONES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          StrSql = "Select * From PROYTAREASADV where codigo = '" & session("NCliente") & "' AND TSK = '" & SESSION("TAREA") & "' ORDER BY id DESC "
          set rsXP = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsXP.eof then


             do while not rsxP.eof

              '  dFecha = cdate(rsXP(3) + clng(cdate("01-01-1900")))
              '  SINPUTn1 = dFecha


                YXTDUR = RSXP(4)
                YXUDUR = RSXP(7)


                ''' RSXP(3) FECHA IMPUTACION
                ''' RSXP(4) TOTAL IMPUTACION
'''''''''''''''''''''''''''''''''''''''''''''''''''' CALCULO HORAS IMPUTADAS ''''''''''''''''''''''''''''''''

               YXCUENTASUM = 0

               SELECT CASE CSTR(YXUDUR & "")


               CASE "2"

                 YXCUENTASUM = YXTDUR * CLng(Session("DPP"))


               CASE "3"

                 YXCUENTASUM = YXTDUR * CLng(Session("DPP")) * CLng(Session("SPP"))


               CASE "4"

                 YXCUENTASUM = REPLACE(CSTR(YXTDUR / 60),",",".")



               CASE ELSE

                 YXCUENTASUM = YXTDUR

               END SELECT


'''''''''''''''''''''''''''''''''''''''''''''''''''' FIN CALCULO HORAS '''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''' ACTUALIZA CUENTA ANTERIOR '''''''''''''''''''''''''''''''''''''''''''''

                YXCUENTADET = RSXP(8)


                strsql = "UPDATE PROYCUENTASDET SET TCONS = (TCONS - " & YXCUENTASUM & ") WHERE ID = 0" & YXCUENTADET
                call ofv.crearconsultaEx(StrSql,Xcn,1,parametros)





''''''''''''''''''''''''''''''''''''' BUSQUEDA Y ACTUALIZACION NUEVA CUENTA ''''''''''''''''''''''''''''''''''



                YXCUENTADET = 0
                YXSTART     = RSXP(3)

                StrSql = "Select ID From PROYcuentasDET WHERE CODCUENTA = 0" & YCUENTA
                strsql = strsql & " AND NOT (FINICIO > " & YXSTART & ") AND NOT (FFINAL < " & YXSTART & ") "
                strsql = strsql & " ORDER BY ACTIVA "
                set rsXP1 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
                if not rsXP1.eof then

                   YXCUENTADET = RSXP1(0)

                ELSE


                   XCODORDEN = CLNG("0")


                   StrSql = "Select MAX(ORDEN) From PROYcuentasDET WHERE CODCUENTA = 0" & YCUENTA
                   set rsXP2 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
                   if not rsXP2.eof then XCODORDEN = RSXP2(0)
                   call ofv.cerrarconsulta(rsXP2)

                   XCODORDEN = XCODORDEN + 10
                   YXTSK     = GTTSK()
                   YXFDHOY   = CLNG(DATE()) - clng(cdate("01-01-1900"))


                   strsql = "insert into PROYCUENTASDET (CODCUENTA,DESCRIPCION,ORDEN,FINICIO,FFINAL,ACTIVA,TPLAN,TCONS) "
                   strsql = strsql & "values ('" & YCUENTA & "','" & YXTSK & "','" & XCODORDEN & "','" & YXSTART & "','99999','1','1','0')"
                   call ofv.crearconsultaEx(StrSql,Xcn,1,parametros)


                   StrSql = "Select ID From PROYcuentasDET WHERE DESCRIPCION = '" & YXTSK & "'"
                   set rsXP2 = ofv.crearconsultaEx(StrSql,Xcn,1,parametros)
                   if not rsXP2.eof then
                      YXCUENTADET = RSXP2(0)
                   END IF
                   call ofv.cerrarconsulta(rsXP2)


               END IF
               call ofv.cerrarconsulta(rsXP1)



               strsql = "UPDATE PROYCUENTASDET SET TCONS = (TCONS + " & YXCUENTASUM & ") WHERE ID = 0" & YXCUENTADET
            '''   RESPONSE.WRITE STRSQL & "<BR>"
               call ofv.crearconsultaEx(StrSql,Xcn,1,parametros)

               strsql = "UPDATE PROYTAREASADV SET CODINPUT = " & YXCUENTADET & " WHERE ID = 0" & RSXP(0)
               call ofv.crearconsultaEx(StrSql,Xcn,1,parametros)



''''''''''''''''''''''' FIN ACTUALIZACION DE IMPUTACION Y CUENTA TRANSFERIDA



             rsxP.movenext
             loop  

          end if
          call ofv.cerrarconsulta(rsXP)



'''''''''' FIN LOOP DE IMPUTACIONES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    StrSql = "UPDATE PROYTAREAS SET TSKCUENTA = " & YCUENTA & " WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
    CALL ofv.crearconsultaEx(StrSql,Xcn,1,parametros)



'''''''''  FIN TRANSFERENCIA DE IMPUTACIONES DE LA TAREA ''''''''''''''''''''''''''''''''''''''''


     END IF
     call ofv.cerrarconsulta(rsX)

        
     Call OFV.cerrarconn(XCN)


     GENTSKCUENTATRF = YXCUENTADET


End Function

''''''''''''''''''''''''''


Public Function GENPROY(YPRY)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

If Session("PRCF") <> "S" Then
   xxxx = GETPRYCNFG()
End If


    If YPRY = "NEW" Then
       CODREQ = GTPROY()

          Yxrtit = "NUEVO PROYECTO"
          Yxrtipo = "0"
          Yxrorig = "0"
          Yxruori = Session("USUARIO")
          Yxruasi = Session("USUARIO")
          Yxrts = "0"
          YXTID = "NO"
          YXRID = "0"
          YXPRI = "3"
    Else

      if wkaplic = "OR" then

       CODREQ = YPRY

       Set CNX2 = OFV.conectar(OFV.strconn2)
       StrSql = "Select * From ORAGENDA WHERE CODAGENDA = '" & YPRY & "' "
       Set RSXP = CrearConsultaEx(StrSql, CNx2, 1, "")
       If Not RSXP.EOF Then
          Yxrtit = "PLAN MITIGACION " & RSXP(1)
          Yxruori = RSXP(6)
          Yxruasi = RSXP(6)
          YXTID = "OR"
          YXRID = RSXP(0)
          YXPRI = "3"
       End If
       Call OFV.cerrarconsulta(RSXP)
       call ofv.cerrarconn(cnx2)

 
          Yxrtipo = "0"
          Yxrorig = "0"
          Yxrts   = "0"

      else

       CODREQ = YPRY

       StrSql = "Select * From REQUERIMIENTO where codigo = '" & YPRY & "' "
       Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
       If Not RSXP.EOF Then
          Yxrtit = "PROYECTO " & RSXP(2)
          Yxrtipo = RSXP(8)
          Yxrorig = RSXP(7)
          Yxruori = RSXP(13)
          Yxruasi = RSXP(14)
          Yxrts = RSXP(9)
          YXTID = "REQ"
          YXRID = RSXP(0)
          YXPRI = RSXP(18)
       End If
     Call OFV.cerrarconsulta(RSXP)


      end if

    End If



If nFecini < CLng(CDate("01-01-1900")) Then
   nFecini = CLng(CDate("01-01-1900"))
Else
   nFecini = 0
End If

vNuevoValor = CLng(Date) - nFecini
    YXFHORA = Right(Now(), Len(Now()) - InStr(Now(), " "))

    StrSql = "insert into PROYECTO (CODIGO,TITULO,FECHAINI,FECHASOL,FECHAENT,FECHAFIN,ORIGEN,TIPO,"
    StrSql = StrSql & "TEMASUBTEMA,DESCRIPCION,ESTADO,DESCTIPO,ORIGINANTE,ASIGNADO,TIPOPROY,OBJPROY,VPLAN,TFECHA,PRIORIDAD,CUENTA,MODAL) "
    StrSql = StrSql & "values ('" & CODREQ & "','" & Yxrtit & "','" & vNuevoValor & "','0','0','0',"
    StrSql = StrSql & "'" & Yxrorig & "','" & Yxrtipo & "','" & Yxrts & "','DESCRIPCION','1','NO',"
    StrSql = StrSql & "'" & Yxruori & "','" & Yxruasi & "','" & YXTID & "','" & YXRID & "','0','" & Session("DHC") & "','" & YXPRI & "','1','1')"
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "INSERT INTO PROYPLAN (CODIGO,VPLAN,DESCRIPCION,FECHAPLAN,HORAPLAN,USRRESP,ESTADO) VALUES ("
    StrSql = StrSql & "'" & CODREQ & "','0','PLAN ORIGINAL','" & vNuevoValor & "','" & YXFHORA & "','" & Yxruasi & "','1')"
    Call CrearConsultaEx(StrSql, CN, 1, "")


  Set fso = CreateObject("Scripting.FileSystemObject")

  IF  YXTID = "REQ" THEN


  ELSE

  if wkaplic <> "GP" then

    sPath = Server.mappath("../CORADMREQ/PROYECTOS")


  ELSE

    sPath = Server.mappath("./PROYECTOS")

  END IF


  Set Folder = fso.getfolder(sPath)
  Set sfolder = Folder.SubFolders
  sfolder.Add (CODREQ)

 END IF

  GENPROY = CODREQ

Call OFV.cerrarconn(CN)
End Function

''''''''''''''''''''''''



Public Function GETTSKCLC(RSXP7, YWXFHOYI, YWXFREALI, YWXFFCHI)

Dim ARRTSK(200, 1)

  For IIII = 0 To RSXP7.Fields.Count - 1
      ARRTSK(IIII, 0) = RSXP7(IIII)
  Next

  XIYTSK = 0

  GETTSKCLC = GETTSKCLCNK(ARRTSK, XIYTSK, YWXFHOYI, YWXFREALI, YWXFFCHI)

End Function

''''''''''''''''''''''''


Public Function GETTSKCLCNK(ARRTSK, XIYTSK, YWXFHOYI, YWXFREALI, YWXFFCHI)


   

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 




    YWXAGRD = 0
    YWXAGRDu = 0
    YWXAGRR = 0
    YWXFHOY = YWXFHOYI
    YWXFREAL = YWXFREALI
    YWXFFDU = 1
    YWXFFTI = 1
    YWXFFCH = YWXFFCHI
    YWPADV = 0
    YWTSKFS = 0
    YWTSKFE = 0
    YWTSKFEW = 0
    YWTSKFSR = 0
    YWTSKFER = 0
    YWTSKST = "0"
    YWXTSKSTART = 0
    YWFSPLAN = 0
    YWFEPLAN = 0
    YWFSREAL = 0
    YWFEREAL = 0



    If ARRTSK(15, XIYTSK) > 0 Then
       YWXFFDU = ARRTSK(15, XIYTSK) 'RSXP7(15)
       YWXFFTI = ARRTSK(16, XIYTSK) 'RSXP7(16)
    Else
       YWXFFDU = 1
       YWXFFTI = 1
    End If

    If ARRTSK(24, XIYTSK) > 0 And (ARRTSK(22, XIYTSK) > 0 And (ARRTSK(22, XIYTSK) <> ARRTSK(15, XIYTSK) Or ARRTSK(23, XIYTSK) <> ARRTSK(16, XIYTSK))) Then
       YWXFFDU = ARRTSK(22, XIYTSK)
       If ARRTSK(23, XIYTSK) > 0 Then YWXFFTI = ARRTSK(23, XIYTSK)
    End If

    YWTSKST = ARRTSK(18, XIYTSK) & ""

    If YWTSKST = "" Then YWTSKST = "0"


    YWTSKFS = ARRTSK(8, XIYTSK) 'RSXP7(8)
    YWTSKFE = ARRTSK(10, XIYTSK) 'RSXP7(10)
    YWTSKFEW = ARRTSK(10, XIYTSK) 'RSXP7(10)
    YWTSKFSR = ARRTSK(9, XIYTSK) 'RSXP7(9)
    YWTSKFER = ARRTSK(11, XIYTSK) 'RSXP7(11)

    If ARRTSK(6, XIYTSK) = 0 And ARRTSK(7, XIYTSK) > 0 And YWTSKST = "0" Then

       YWXTSKSTART = GETFSTARTV(ARRTSK(3, XIYTSK), ARRTSK(1, XIYTSK))
       If YWXTSKSTART > 0 Then
          YWXFREAL = YWXTSKSTART
          YWTSKFSR = YWXFREAL
       End If



    End If


    If ARRTSK(6, XIYTSK) = 0 And ARRTSK(7, XIYTSK) = 0 And YWTSKST = "0" Then


       if YWTSKFS  = 0 then
          YWTSKFS  = clng(date()) - clng(cdate("01-01-1900"))
       end if 


       YWXTSKSTART = YWTSKFS
       YWXFREAL    = YWXTSKSTART
       YWTSKFSR    = YWXFREAL

       YWXAGRW     = GETFECHA(YWTSKFS, YWXFFDU, YWXFFTI, YWXFFCH)
       YWTSKFE     = WTFCH(3)
       YWTSKFER    = YWTSKFE


    End If



     If YWTSKFS = 0 Or YWTSKFS = 11111 Then
        If YWTSKFSR > 0 Then
           YWTSKFS = YWTSKFSR

           YWXAGRW = GETFECHA(YWTSKFS, YWXFFDU, YWXFFTI, YWXFFCH)
           YWTSKFE = WTFCH(3)
           YWTSKFER = YWTSKFE

        Else
           YWTSKFS = YWXFREAL
           YWTSKFSR = YWXFREAL

           YWXAGRW = GETFECHA(YWTSKFS, YWXFFDU, YWXFFTI, YWXFFCH)
           YWTSKFE = WTFCH(3)
           YWTSKFER = YWTSKFE

        End If
    End If

    YWXAGRW = GETFECHA(YWTSKFS, YWXFFDU, YWXFFTI, YWXFFCH)

    If YWTSKFE = 0 Then
         YWTSKFE = WTFCH(3)
    End If

       YWXAGRP = WTFCH(1)
       YWFSPLAN = YWTSKFS
       YWFEPLAN = YWTSKFE

   YWSTSSIGN = 0

   If YWTSKST = "0" Then

        YWFSREAL = YWXFREAL
        YWTSKFSR = YWXFREAL

        YWXAGRW = GETFECHA(YWXFREAL, YWXFFDU, YWXFFTI, YWXFFCH)
        YWXFREAL = WTFCH(3)

        YWFEREAL = YWXFREAL
        YWTSKFER = YWXFREAL

        YWXAGRR = WTFCH(1)
        YWXAGRD = WTFCH(1)








   ElseIf YWTSKST = "2" Then



         YWFSREAL = YWTSKFSR
         YWFEREAL = YWTSKFER
         YWXFREAL = YWTSKFER

         YWXAGRW = GETFECHA(YWXFREAL, YWXFFDU, YWXFFTI, YWXFFCH)

         YWXAGRR = YWFEREAL - YWFSREAL
         YWXAGRD = WTFCH(1)







   Else


          YWXFREAL = YWTSKFSR

          YWPADV = 0

          Set CN = OFV.conectar(OFV.strconn6)

           TSKADV = "N"

          StrSql = "Select * From PROYTAREASADV where codigo = '" & ARRTSK(1, XIYTSK) & "' AND TSK = '" & ARRTSK(3, XIYTSK) & "' ORDER BY id desc "
          Set RSXP8 = CrearConsultaEx(StrSql, CN, 1, "")
          If Not RSXP8.EOF Then

             YWPADV = RSXP8(5)
              TSKADV = "S"

          End If
          Call OFV.cerrarconsulta(RSXP8)
          Call OFV.cerrarconn(CN)

           If TSKADV = "S" Then

             If YWXFHOY < YWTSKFE Then


                 YWSTSSIGN = 0

                YWXAGRW = GETFECHA(YWXFREAL, YWXFFDU, YWXFFTI, YWXFFCH)
                YWXFREAL = WTFCH(3)

                YWXAGRR = WTFCH(1)
                YWXAGRD = WTFCH(1)

                 ZXCONS = (WTFCH(0) * YWPADV / 100)
                 ZXREST = CInt(((WTFCH(0) - ZXCONS) * CLng(Session("DPP"))) / CLng(Session("MPP")))
                 ZXAGRW = GETFECHA(YWXFHOY, ZXREST, 2, YWXFFCH)

                 ZXFREAL = WTFCH(3)

                 If ZXFREAL > YWTSKFE Then
                    YWSTSSIGN = 2

                    YWXAGRW = GETFECHA(YWXFHOY, YWXFFDU, YWXFFTI, YWXFFCH)
                    YWXAGRDu = WTFCH(1)

                    YWXAGRW = YWXFFDU - CInt((YWXFFDU * YWPADV / 100))
                    YWXAGRW = GETFECHA(YWXFHOY, YWXAGRW, YWXFFTI, YWXFFCH)
                    YWXAGRR = (YWXFHOY - YWXFREAL) + WTFCH(1)
                    YWXFREAL = WTFCH(3)


                    YWXAGRD = YWXAGRDu
     
                 End If

             ElseIf YWXFHOY = YWTSKFE And YWPADV > 99 Then

                 YWSTSSIGN = 0

                YWXAGRW = GETFECHA(YWXFREAL, YWXFFDU, YWXFFTI, YWXFFCH)
                YWXFREAL = WTFCH(3)

                YWXAGRR = WTFCH(1)
                YWXAGRD = WTFCH(1)


             Else

                 YWSTSSIGN = 2


                YWXAGRW = GETFECHA(YWXFHOY, YWXFFDU, YWXFFTI, YWXFFCH)
                YWXAGRDu = WTFCH(1)

                YWXAGRW = YWXFFDU - CInt((YWXFFDU * YWPADV / 100))
                YWXAGRW = GETFECHA(YWXFHOY, YWXAGRW, YWXFFTI, YWXFFCH)
                YWXAGRR = (YWXFHOY - YWXFREAL) + WTFCH(1)
                YWXFREAL = WTFCH(3)


                YWXAGRD = YWXAGRDu





             End If

          Else

            If (YWXFHOY < YWTSKFE) Or (YWXFHOY = YWTSKFE) Then

                YWSTSSIGN = 1

               YWXAGRW = GETFECHA(YWXFREAL, YWXFFDU, YWXFFTI, YWXFFCH)
               YWXFREAL = WTFCH(3)

               YWXAGRR = WTFCH(1)
               YWXAGRD = WTFCH(1)

                 ZXCONS = WTFCH(0)
                 ZXREST = CInt(((ZXCONS) * CLng(Session("DPP"))) / CLng(Session("MPP")))
                 ZXAGRW = GETFECHA(YWXFHOY, ZXREST, 2, YWXFFCH)


                 ZXFREAL = WTFCH(3)

                 If ZXFREAL > YWTSKFE Then
                    YWSTSSIGN = 2

                    YWXAGRW = GETFECHA(YWXFHOY, YWXFFDU, YWXFFTI, YWXFFCH)
                    YWXAGRDu = WTFCH(1)

                    YWXAGRW = YWXFFDU - CInt((YWXFFDU * YWPADV / 100))
                    YWXAGRW = GETFECHA(YWXFHOY, YWXAGRW, YWXFFTI, YWXFFCH)
                    YWXAGRR = (YWXFHOY - YWXFREAL) + WTFCH(1)
                    YWXFREAL = WTFCH(3)


                    YWXAGRD = YWXAGRDu
     
                 End If

            Else

                YWSTSSIGN = 2
              


                    YWXAGRW = GETFECHA(YWXFHOY, YWXFFDU, YWXFFTI, YWXFFCH)
                    YWXAGRDu = WTFCH(1)

                    YWXAGRW = YWXFFDU - CInt((YWXFFDU * YWPADV / 100))
                    YWXAGRW = GETFECHA(YWXFHOY, YWXAGRW, YWXFFTI, YWXFFCH)
                    YWXAGRR = (YWXFHOY - YWXFREAL) + WTFCH(1)
                    YWXFREAL = WTFCH(3)


                    YWXAGRD = YWXAGRDu


            End If



          End If


        YWFSREAL = YWTSKFSR
        YWFEREAL = YWXFREAL
        YWTSKFER = YWXFREAL
        If YWTSKFEW = 0 Then
           YWTSKFE = YWXFREAL
           YWFEPLAN = YWTSKFE
           YWXAGRP = YWXAGRD
        End If




   End If


   WIIX = GETTABTSK()


   ATABTSKW(WIIX, 0) = ARRTSK(3, XIYTSK)
   ATABTSKW(WIIX, 1) = YWXFREAL
   ATABTSKW(WIIX, 2) = YWTSKFS
   ATABTSKW(WIIX, 3) = YWTSKFE
   ATABTSKW(WIIX, 4) = YWTSKFSR
   ATABTSKW(WIIX, 5) = YWTSKFER
   ATABTSKW(WIIX, 6) = YWTSKST
   ATABTSKW(WIIX, 7) = YWFSPLAN
   ATABTSKW(WIIX, 8) = YWFEPLAN
   ATABTSKW(WIIX, 9) = YWFSREAL
   ATABTSKW(WIIX, 10) = YWFEREAL
   ATABTSKW(WIIX, 11) = YWXFFDU
   ATABTSKW(WIIX, 12) = YWXFFTI
   ATABTSKW(WIIX, 13) = YWPADV
   ATABTSKW(WIIX, 14) = YWXAGRD
   ATABTSKW(WIIX, 15) = YWXAGRR
   ATABTSKW(WIIX, 16) = ARRTSK(24, XIYTSK)
   ATABTSKW(WIIX, 17) = ARRTSK(15, XIYTSK)
   ATABTSKW(WIIX, 18) = ARRTSK(16, XIYTSK)
   ATABTSKW(WIIX, 19) = ARRTSK(22, XIYTSK)
   ATABTSKW(WIIX, 20) = ARRTSK(23, XIYTSK)
   ATABTSKW(WIIX, 21) = YWXAGRP
   ATABTSKW(WIIX, 22) = YWSTSSIGN

   GETTSKCLCNK = WIIX
   


End Function

'''''''''''''''''''''''''
Public Function PRYPLANRVRS(YPRY, YDESCRI, YVPLANV, YVPLANN, YUSR, YXFPRYSTRT, YLNCH)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

  If YLNCH = "R" Or YLNCH = "C" Or YLNCH = "KC" Then

    StrSql = "DELETE FROM PROYETAPA WHERE  CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "DELETE FROM PROYTAREAS WHERE  CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")


  End If

  If YLNCH = "R" Then

    StrSql = "DELETE FROM PROYPLAN WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYECTO SET VPLAN = " & YVPLANV & ", ESTADO = 1,  FECHAENT = 0 ,  FECHASOL = " & YXFPRYSTRT & ",DESCRIPCION = '" & YDESCRI & "' WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYPLAN SET ESTADO = 1 WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANV
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYTAREAS SET FECHASPLAN = 0, FECHAEPLAN = 0 WHERE CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANV
    Call CrearConsultaEx(StrSql, CN, 1, "")

 End If

  If YLNCH = "C" Then

    StrSql = "UPDATE PROYPLAN SET ESTADO = 3 WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYECTO SET ESTADO = 12 WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN, 1, "")

 End If

  If YLNCH = "P" Then

    StrSql = "UPDATE PROYPLAN SET ESTADO = 2 WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANV
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYPLAN SET ESTADO = 1 WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYECTO SET ESTADO = 12, VPLAN = " & YVPLANN & " WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN, 1, "")

 End If

  If YLNCH = "K" Then

    StrSql = "DELETE FROM PROYETAPA WHERE  CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANV
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "DELETE FROM PROYTAREAS WHERE  CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANV
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYETAPA SET VPLAN = " & YVPLANV & " WHERE  CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYTAREAS SET VPLAN = " & YVPLANV & " WHERE  CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")


  End If

  If YLNCH = "K" Or YLNCH = "KC" Then



    StrSql = "DELETE FROM PROYPLAN WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYECTO SET ESTADO = 12 WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN, 1, "")

 End If
 
 Call OFV.cerrarconn(CN)

End Function


'''''''''''''''''''''''''

'''''''''''''''''''''''''
Public Function PRYPLAN(YPRY, YDESCRI, YVPLANV, YVPLANN, YUSR, YXFPRYSTRT, YLNCH)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

  If YVPLANN = "NEXT" Then
     StrSql = "Select MAX(VPLAN) FROM PROYPLAN where codigo = '" & YPRY & "' "
     Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
     If Not RSXP.EOF Then
        YVPLANN = RSXP(0) + 1
     Else
        YVPLANN = 2
     End If
     Call OFV.cerrarconsulta(RSXP)
  End If

    YXFHORA = Right(Now(), Len(Now()) - InStr(Now(), " "))

    StrSql = "INSERT INTO PROYPLAN (CODIGO,VPLAN,DESCRIPCION,FECHAPLAN,HORAPLAN,USRRESP,ESTADO) VALUES ("
    StrSql = StrSql & "'" & YPRY & "','" & YVPLANN & "','" & YDESCRI & "','" & YXFPRYSTRT & "','" & YXFHORA & "','" & YUSR & "','0')"
    Call CrearConsultaEx(StrSql, CN, 1, "")


  StrSql = "Select * From PROYetapa where codigo = '" & YPRY & "' AND VPLAN = " & YVPLANV & " order by orden "
  Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
  If Not RSXP.EOF Then
     Do While Not RSXP.EOF


        StrSql = "Insert into PROYetapa ("
        strsql2 = ") values ("

    For ii = 1 To (RSXP.Fields.Count - 1)
      If ii > 1 Then
         StrSql = StrSql & ","
         strsql2 = strsql2 & ","
      End If

      StrSql = StrSql & RSXP(ii).Name

      s = "'"
      If UCase(RSXP(ii).Name) = "VPLAN" Then
         sInput = YVPLANN
      Else
         sInput = RSXP(ii) & ""
      End If

      If UCase(sInput) = "VERDADERO" Then sInput = "1"
      If UCase(sInput) = "FALSO" Then sInput = "0"
      If UCase(sInput) = "TRUE" Then sInput = "1"
      If UCase(sInput) = "FALSE" Then sInput = "0"
 

      If RSXP(ii).Type = 131 Then sInput = Replace(RSXP(ii), ",", ".")

   
      strsql2 = strsql2 & s & sInput & s

    Next
  
    StrSql = StrSql & strsql2 & ") "
    Call CrearConsultaEx(StrSql, CN, 1, "")
     
     RSXP.MoveNext
     Loop

  End If

  Call OFV.cerrarconsulta(RSXP)

  YFECSHOY = CLng(Date) - CLng(CDate("01-01-1900"))
  YSSUMA = 0


  StrSql = "Select * From PROYtareas where codigo = '" & YPRY & "' and VPLAN = " & YVPLANV & " order by orden "
  Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
  If Not RSXP.EOF Then
     Do While Not RSXP.EOF

     If RSXP(4) = 10 Then YSSUMA = 0



     If RSXP(18) = 1 Then

        SXFSTARTD = GETFENDV(RSXP(3), RSXP(1))
        XIIX = TIX

XFREAL = ATABTSKW(XIIX, 1)
TSKFS = ATABTSKW(XIIX, 2)
TSKFE = ATABTSKW(XIIX, 3)
TSKFSR = ATABTSKW(XIIX, 4)
TSKFER = ATABTSKW(XIIX, 5)
TSKST = ATABTSKW(XIIX, 6)
FSPLAN = ATABTSKW(XIIX, 7)
FEPLAN = ATABTSKW(XIIX, 8)
FSREAL = ATABTSKW(XIIX, 9)
FEREAL = ATABTSKW(XIIX, 10)
XFFDU = ATABTSKW(XIIX, 11)
XFFTI = ATABTSKW(XIIX, 12)
PADV = ATABTSKW(XIIX, 13)
XAGRD = ATABTSKW(XIIX, 14)
XAGRR = ATABTSKW(XIIX, 15)
XMODA = ATABTSKW(XIIX, 16)
XTDUR = ATABTSKW(XIIX, 17)
XUDUR = ATABTSKW(XIIX, 18)
XTDUP = ATABTSKW(XIIX, 19)
XUDUP = ATABTSKW(XIIX, 20)
XAGRP = ATABTSKW(XIIX, 21)

XFSPLAN = FSPLAN
XFEPLAN = FEPLAN
XFSREAL = FSREAL
XFEREAL = FEREAL

    
If FEPLAN < FEREAL Then

   If PADV = 0 Then
      YSMARCA = "SN"

      XFSPLAN = YFECSHOY
      XFEPLAN = 0
      XFSREAL = YFECSHOY
      XFEREAL = 0

   Else
      YSMARCA = "SS"

      XFSPLAN = YFECSHOY
      XFEPLAN = 0
      XFSREAL = YFECSHOY
      XFEREAL = 0

      Select Case XUDUR
                        Case 4
                YXAGRD = XTDUR
                YxagrT = 4

           Case 1
                YXAGRD = XTDUR
                YxagrT = 1

           Case 2

                YXAGRD = (XTDUR * CLng(Session("DPP")))
                YxagrT = 1

           Case Else

                YXAGRD = (XTDUR * CLng(Session("DPP")) * CLng(Session("SPP")))
                YxagrT = 1

      End Select


      YWXAGRW = CInt((YXAGRD * PADV / 100))

      FTDUR = YWXAGRW
      FUDUR = YxagrT
      FTDUP = 0
      FUDUP = 0
      FFTPT = 0

      XTDUR = YXAGRD - YWXAGRW
      XUDUR = YxagrT

      If RSXP(24) > 0 Then

      Select Case XUDUP
                        Case 4
                YXAGRD = XTDUP
                YxagrT = 4

           Case 1
                YXAGRD = XTDUP
                YxagrT = 1

           Case 2

                YXAGRD = (XTDUP * CLng(Session("DPP")))
                YxagrT = 1

           Case Else

                YXAGRD = (XTDUP * CLng(Session("DPP")) * CLng(Session("SPP")))
                YxagrT = 1

      End Select


         YWXAGRW2 = CInt((YXAGRD * PADV / 100))
         XTDUP = YXAGRD - YWXAGRW2
         XUDUP = YxagrT
         XFTPT = 1
      Else
         XTDUP = 0
         XUDUP = 0
         XFTPT = 0
      End If

      FFSPLAN = FSPLAN
      FFEPLAN = YFECSHOY
      FFSREAL = FSREAL
      FFEREAL = YFECSHOY



   End If


Else

        YSMARCA = "N"

End If


     Else

        YSMARCA = "N"

     End If

     StrSql = "Insert into PROYtareas ("
     strsql2 = ") values ("

    For ii = 1 To (RSXP.Fields.Count - 1)
      If ii > 1 Then
         StrSql = StrSql & ","
         strsql2 = strsql2 & ","
      End If

      StrSql = StrSql & RSXP(ii).Name

      s = "'"
      If UCase(RSXP(ii).Name) = "VPLAN" Then
         sInput = YVPLANN
      Else
         sInput = RSXP(ii) & ""
      End If

   If UCase(RSXP(ii).Name) = "ORDEN" Then
      sInput = RSXP(ii) + YSSUMA
   End If


      If UCase(sInput) = "VERDADERO" Then sInput = "1"
      If UCase(sInput) = "FALSO" Then sInput = "0"
      If UCase(sInput) = "TRUE" Then sInput = "1"
      If UCase(sInput) = "FALSE" Then sInput = "0"
 

      If RSXP(ii).Type = 131 Then sInput = Replace(RSXP(ii), ",", ".")

   If YSMARCA = "SN" Then

      Select Case UCase(RSXP(ii).Name)

                   Case "FECHASPLAN"

            sInput = XFSPLAN

      Case "FECHAEPLAN"

            sInput = XFEPLAN

      Case "FECHASREAL"

            sInput = XFSREAL

      Case "FECHAEREAL"

            sInput = XFEREAL


      End Select

   End If

   If YSMARCA = "SS" Then
      Select Case UCase(RSXP(ii).Name)

                   Case "TSKSPL"

            sInput = RSXP("TAREA") & "VP" & YVPLANN

      Case "FECHASPLAN"

            sInput = FFSPLAN

      Case "FECHAEPLAN"

            sInput = FFEPLAN

      Case "FECHASREAL"

            sInput = FFSREAL

      Case "FECHAEREAL"

            sInput = FFEREAL

      Case "TDUR"

            sInput = FTDUR

      Case "UDUR"

            sInput = FUDUR

      Case "CASIG"

            sInput = FTDUP

      Case "TASIG"

            sInput = FUDUP

      Case "MASIG"

            sInput = FFTPT

      Case "ESTADO"

            sInput = 2


      End Select


   End If


   
      strsql2 = strsql2 & s & sInput & s

    Next
  
    STRSQLX = StrSql & strsql2 & ") "
    Call CrearConsultaEx(STRSQLX, CN, 1, "")

 If YSMARCA = "SS" Then

    YSSUMA = YSSUMA + 10
    strsql2 = ") values ("

  For ii = 1 To (RSXP.Fields.Count - 1)
   If ii > 1 Then
      strsql2 = strsql2 & ","
   End If

   s = "'"
   If UCase(RSXP(ii).Name) = "VPLAN" Then
      sInput = YVPLANN
   Else
      sInput = RSXP(ii) & ""
   End If

   If UCase(RSXP(ii).Name) = "ORDEN" Then
      sInput = RSXP(ii) + YSSUMA
   End If

   If UCase(sInput) = "VERDADERO" Then sInput = "1"
   If UCase(sInput) = "FALSO" Then sInput = "0"
   If UCase(sInput) = "TRUE" Then sInput = "1"
   If UCase(sInput) = "FALSE" Then sInput = "0"
 

   If RSXP(ii).Type = 131 Then sInput = Replace(RSXP(ii), ",", ".")


      Select Case UCase(RSXP(ii).Name)



                   Case "TAREA"

            sInput = RSXP(ii) & "VP" & YVPLANN

      Case "DESCRIPCION"

            sInput = RSXP(ii) & " (SP)"

      Case "FECHASPLAN"

            sInput = XFSPLAN

      Case "FECHAEPLAN"

            sInput = XFEPLAN

      Case "FECHASREAL"

            sInput = XFSREAL

      Case "FECHAEREAL"

            sInput = XFEREAL

      Case "TDUR"

            sInput = XTDUR

      Case "UDUR"

            sInput = XUDUR

      Case "CASIG"

            sInput = XTDUP

      Case "TASIG"

            sInput = XUDUP

      Case "MASIG"

            sInput = XFTPT

      Case "ESTADO"

            sInput = 1


      End Select





   
    strsql2 = strsql2 & s & sInput & s

   Next
  
   STRSQLX = StrSql & strsql2 & ") "
       
   Call CrearConsultaEx(STRSQLX, CN, 1, "")

 End If

     
     RSXP.MoveNext
     Loop
 
  End If

  Call OFV.cerrarconsulta(RSXP)





  If YLNCH = "S" Then



    StrSql = "UPDATE PROYPLAN SET ESTADO = 2 WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANV
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYPLAN SET ESTADO = 1 WHERE CODIGO = '" & YPRY & "'  AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")


    StrSql = "UPDATE PROYECTO SET VPLAN = " & YVPLANN & " WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN, 1, "")

 End If

  If YLNCH = "P" Then



    StrSql = "UPDATE PROYTAREAS SET FECHASPLAN = 0, FECHAEPLAN = 0 WHERE ESTADO = 0 AND CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYTAREAS SET FECHASPLAN = FECHASREAL ,FECHAEPLAN = 0 WHERE ESTADO = 1 AND CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYECTO SET ESTADO = 15 WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN, 1, "")

 End If

  If YLNCH = "K" Then



    StrSql = "UPDATE PROYTAREAS SET FECHASPLAN = 0, FECHAEPLAN = 0 WHERE ESTADO = 0 AND CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYTAREAS SET FECHASPLAN = FECHASREAL ,FECHAEPLAN = 0 WHERE ESTADO = 1 AND CODIGO = '" & YPRY & "' AND VPLAN = " & YVPLANN
    Call CrearConsultaEx(StrSql, CN, 1, "")

    StrSql = "UPDATE PROYECTO SET ESTADO = 16 WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN, 1, "")

 End If

    PRYPLAN = YVPLANN
    
    Call OFV.cerrarconn(CN)

End Function

'''''''''''''''''''''''''

Public Function PRYLNCH(YPRY, YDESCRI, YLNCH)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN3 = OFV.conectar(OFV.strconn6)

    YXFSTART = CLng(Date) - CLng(CDate("01-01-1900"))
    YXFCURR = YXFSTART




  If YLNCH <> "P" And YLNCH <> "K" Then

    xxxx = PRYPLANRVRS(YPRY, YDESCRI, "0", "1", "NO", YXFSTART, "R")

  End If




     YNPLAN = Session("VPLAN")


    StrSql = "Select * From PROYecto where codigo = '" & YPRY & "'  "
    Set RSXP = CrearConsultaEx(StrSql, CN3, 1, "")
    If Not RSXP.EOF Then
          YXASIG = RSXP(14)
          YXFSTS = RSXP(11)
          YXFFCH = RSXP(18)
          YVPLAN = RSXP(17)
          If CLng(YXFFCH) > 0 Then XXX = ADMFERIADOS("C", 0)
    End If

    If YPRY <> Session("NCliente") And YLNCH <> "K" Then
       YNPLAN = YVPLAN
    End If


  If YLNCH = "P" Or YLNCH = "K" Then


     StrSql = "Select TAREA,TSKSPL,ID From PROYtareas where codigo = '" & YPRY & "' and VPLAN = " & YNPLAN & " AND TSKSPL <> 'NO' order by ID "
     Set RSXP1 = CrearConsultaEx(StrSql, CN3, 1, "")
     If Not RSXP1.EOF Then
        Do While Not RSXP1.EOF
           XNVTAREA = RSXP1(0)
           XNVTSKSPL = RSXP1(1)
           StrSql = "UPDATE PROYtareas SET TSKSPL = 'NO' where ID = " & RSXP1(2)
           Call CrearConsultaEx(StrSql, CN3, 1, "")

           StrSql = "UPDATE PROYtareasLST SET DTAREA = '" & XNVTSKSPL & "' where DTAREA = '" & XNVTAREA & "' AND DCODIGO = '" & YPRY & "' "
           Call CrearConsultaEx(StrSql, CN3, 1, "")

        RSXP1.MoveNext
        Loop


     End If
     Call OFV.cerrarconsulta(RSXP1)

  End If



 YWETP = 0
 YWTSK = 0
          YXFREAL = YXFCURR

  StrSql = "Select * From PROYetapa where codigo = '" & YPRY & "' AND VPLAN = " & YNPLAN & " order by orden "
  Set RSXP = CrearConsultaEx(StrSql, CN3, 1, "")
  If Not RSXP.EOF Then
     Do While Not RSXP.EOF

     YWETP = YWETP + 1

     YXFINIETP = YXFCURR

  StrSql = "Select * From PROYtareas where codigo = '" & YPRY & "' and etapa = '" & RSXP(2) & "' AND VPLAN = " & YNPLAN & " order by orden "
  Set RSXP1 = CrearConsultaEx(StrSql, CN3, 1, "")
  If Not RSXP1.EOF Then
     Do While Not RSXP1.EOF




    XIIX = GETTSKCLC(RSXP1, YXFSTART, YXFREAL, YXFFCH)




    YXFREAL = ATABTSKW(XIIX, 1)
    YTSKFS = ATABTSKW(XIIX, 2)
    YTSKFE = ATABTSKW(XIIX, 3)
    YTSKFSR = ATABTSKW(XIIX, 4)
    YTSKFER = ATABTSKW(XIIX, 5)
    YTSKST = ATABTSKW(XIIX, 6)
    YFSPLAN = ATABTSKW(XIIX, 7)
    YFEPLAN = ATABTSKW(XIIX, 8)
    YFSREAL = ATABTSKW(XIIX, 9)
    YFEREAL = ATABTSKW(XIIX, 10)
    YXFFDU = ATABTSKW(XIIX, 11)
    YXFFTI = ATABTSKW(XIIX, 12)
    YPADV = ATABTSKW(XIIX, 13)
    YXAGRD = ATABTSKW(XIIX, 14)
    YXAGRR = ATABTSKW(XIIX, 15)
    YXMODA = ATABTSKW(XIIX, 16)
    YXTDUR = ATABTSKW(XIIX, 17)
    YXUDUR = ATABTSKW(XIIX, 18)
    YXTDUP = ATABTSKW(XIIX, 19)
    YXUDUP = ATABTSKW(XIIX, 20)








    YXFINI = YTSKFS
    YXFCURR = YXFREAL

  If YLNCH = "P" Or YLNCH = "K" Then
     YXESTADO = RSXP1(18)
  Else
     YXESTADO = "0"
  End If

    YWTSK = YWTSK + 1
    If YWETP = 1 And YWTSK = 1 Then
       YXFPRYSTRT = YXFINI
       YXFPRYFTSK = RSXP1(3)
    End If

    StrSql = "UPDATE PROYTAREAS SET FECHASPLAN = " & YXFINI & ",FECHAEPLAN = " & YXFREAL & ",ESTADO = " & YXESTADO & " WHERE ID = " & RSXP1(0)
     Call CrearConsultaEx(StrSql, CN3, 1, "")


    


    RSXP1.MoveNext
    Loop
  Call OFV.cerrarconsulta(RSXP1)

  End If

   YXFENDETP = YXFCURR

   

    StrSql = "UPDATE PROYETAPA SET FECHASPLAN = " & YXFINIETP & ",FECHAEPLAN = " & YXFREAL & " WHERE ID = " & RSXP(0)
    Call CrearConsultaEx(StrSql, CN3, 1, "")

    RSXP.MoveNext
    Loop
  Call OFV.cerrarconsulta(RSXP)

  End If

  If YLNCH = "P" Or YLNCH = "K" Then

    xxxx = PRYPLANRVRS(YPRY, "ACTIVACION", YVPLAN, YNPLAN, Session("USUARIO"), YXFSTART, YLNCH)
       
  Else

    YNVPLAN = PRYPLAN(YPRY, "PLAN INICIAL", "0", "1", YXASIG, YXFPRYSTRT, "S")

    If YLNCH = "S" Then


       YXESTADO = "1,FECHASREAL = " & YXFPRYSTRT

       StrSql = "UPDATE PROYTAREAS SET ESTADO = " & YXESTADO & " WHERE TAREA = '" & YXFPRYFTSK & "' AND VPLAN = 1 "
       Call CrearConsultaEx(StrSql, CN3, 1, "")

    End If

    StrSql = "UPDATE PROYECTO SET FECHAENT = " & YXFCURR & ",ESTADO = 12 WHERE CODIGO = '" & YPRY & "'  "
    Call CrearConsultaEx(StrSql, CN3, 1, "")

  End If

   
   Call OFV.cerrarconn(CN3)


End Function
'''''''''''''''''''''''''

Public Function TSKLNCHERSV(YPRY, YTSK, YOPC)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

    YXFSTART = CLng(Date) - CLng(CDate("01-01-1900"))
    YTSKCD = YTSK


    If YPRY = Session("NCliente") Then
       YXVPLAN = Session("VPLAN")

    Else

       StrSql = "Select * From PROYecto where codigo = '" & YPRY & "'  "
       Set RSXP2 = CrearConsultaEx(StrSql, CN, 1, "")
       If Not RSXP2.EOF Then
          YXVPLAN = RSXP2(17)
       End If
       Call OFV.cerrarconsulta(RSXP2)
    End If



                  StrSql = "UPDATE PROYTAREAS SET ESTADO = 1, FECHASREAL = " & YXFSTART & " WHERE codigo = '" & YPRY & "' AND TAREA = '" & YTSK & "' AND VPLAN = " & YXVPLAN
                  Call CrearConsultaEx(StrSql, CN, 1, "")

     TSKLNCHERSV = YTSK

Call OFV.cerrarconn(CN)

End Function


Public Function TSKLNCHER(YPADV, YPRY, YTSK, YOPC)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

    YXFSTART = CLng(Date) - CLng(CDate("01-01-1900"))
    YTSKCD = YTSK

    If YPRY = Session("NCliente") Then
       YXVPLAN = Session("VPLAN")

    Else

       StrSql = "Select * From PROYecto where codigo = '" & YPRY & "'  "
       Set RSXP2 = CrearConsultaEx(StrSql, CN, 1, "")
       If Not RSXP2.EOF Then
          YXVPLAN = RSXP2(17)
       End If
       Call OFV.cerrarconsulta(RSXP2)
    End If

    If YPADV > 99 Then
      If YOPC = "S" Then
       StrSql = "UPDATE PROYTAREAS SET ESTADO = 2, FECHAEREAL = " & YXFSTART & " WHERE codigo = '" & YPRY & "' AND TAREA = '" & YTSK & "' AND VPLAN = " & YXVPLAN
       Call CrearConsultaEx(StrSql, CN, 1, "")
      End If




       YPOSTSK = 0

       StrSql = "Select * From PROYETPTSK where codigo = '" & YPRY & "' AND VPLAN = " & YXVPLAN & "  ORDER BY ORDEN,TSKORD,FECHASPLAN DESC "
       Set RSXP5 = CrearConsultaEx(StrSql, CN, 1, "")
       If Not RSXP5.EOF Then
         Do While Not RSXP5.EOF
            If YPOSTSK > 0 Then

               If RSXP5(4) = 0 And RSXP5(6) = 1 Then
                  StrSql = "UPDATE PROYTAREAS SET ESTADO = 1, FECHASREAL = " & YXFSTART & " WHERE codigo = '" & YPRY & "' AND VPLAN = " & YXVPLAN & "  AND ID = " & RSXP5(2)
                  Call CrearConsultaEx(StrSql, CN, 1, "")
                  Exit Do
               Else
                  Exit Do
               End If

            End If

            If RSXP5(20) = YTSK Then
               YPOSTSK = 1
            End If

 

         RSXP5.MoveNext
         Loop

       Else


       End If
       Call OFV.cerrarconsulta(RSXP5)

       YXXX = TSKLNCHLST(YPRY, YTSKCD, YOPC)



    End If
    
    Call OFV.cerrarconn(CN)
    
End Function

Public Function TSKLNCHLST(YPRY, YTSK, YOPC)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

  YXFSTART = CLng(Date) - CLng(CDate("01-01-1900"))

  StrSql = "Select * From PROYTAREASLST where DCODIGO = '" & YPRY & "' and DTAREA = '" & YTSK & "' AND ESTADO = 0 order by orden "
  Set RSXP2 = CrearConsultaEx(StrSql, CN, 1, "")
  Do While Not RSXP2.EOF




     StrSql = "Select * From PROYTAREASLST where CODIGO = '" & RSXP2(1) & "' and TAREA = '" & RSXP2(2) & "' and DTAREA <> '" & YTSK & "' AND ESTADO = 0 order by orden "
     Set RSXP3 = CrearConsultaEx(StrSql, CN, 1, "")
     If RSXP3.EOF Then

        If Not GETPRYSTS(RSXP2(1)) Then

           xxxx = PRYLNCH(RSXP2(1), "SE INICIA EL PROYECTO EN FUNCION DE LA FINALIZACION DE LAS TAREAS PRESEDENTES", "N")
           
        End If

        If RSXP2(1) = Session("NCliente") Then
           YXVPLAN = Session("VPLAN")

        Else

           StrSql = "Select * From PROYecto where codigo = '" & RSXP2(1) & "'  "
           Set RSXP4 = CrearConsultaEx(StrSql, CN, 1, "")
           If Not RSXP4.EOF Then
              YXVPLAN = RSXP4(17)
           End If
           Call OFV.cerrarconsulta(RSXP4)
        End If

        StrSql = "UPDATE PROYTAREAS SET ESTADO = 1, FECHASREAL = " & YXFSTART & " WHERE codigo = '" & RSXP2(1) & "' AND TAREA = '" & RSXP2(2) & "' AND VPLAN = " & YXVPLAN
        Call CrearConsultaEx(StrSql, CN, 1, "")

     End If
     Call OFV.cerrarconsulta(RSXP3)


     StrSql = "UPDATE PROYTAREASLST SET ESTADO = 1 WHERE ID = " & RSXP2(0)
     Call CrearConsultaEx(StrSql, CN, 1, "")

  RSXP2.MoveNext
  Loop
  
  Call OFV.cerrarconsulta(RSXP2)
  
  Call OFV.cerrarconn(CN)


End Function

Public Function GETPRYSTS(YPRY)

Dim RSX 
Dim RSXP 
Dim RSXP1 
Dim RSXP2 
Dim RSXP3 
Dim RSXP4 
Dim RSXP5 
Dim RSXP6 
Dim RSXP7 
Dim RSXP8 

Dim CN 
Dim CN2 
Dim CN3 
Dim CN4 
Dim CN5 


Set CN = OFV.conectar(OFV.strconn6)

    GETPRYSTS = (0 > 1)

    StrSql = "Select * From PROYECTO where CODIGO = '" & YPRY & "'  "
    Set RSXP = CrearConsultaEx(StrSql, CN, 1, "")
    If Not RSXP.EOF Then
       GETPRYSTS = (RSXP(11) > 1)
    End If
    
    Call OFV.cerrarconn(CN)

End Function



'''''''''''''''''''''''''

Public Function GTWKF()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "WKF"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTWKF = Trim(GTsecuf)

End Function

Public Function GTWKFPSC()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "WKFPSC"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTWKFPSC = Trim(GTsecuf)

End Function

Public Function GTCKT()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "CKT"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTCKT = Trim(GTsecuf)

End Function

Public Function GTCKTPSC()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "CKTPSC"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTCKTPSC = Trim(GTsecuf)

End Function

Public Function GTREQ()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "REQ"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTREQ = Trim(GTsecuf)

End Function

Public Function GTPROY()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "PRY"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTPROY = Trim(GTsecuf)

End Function

Public Function GTETP()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "ETP"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTETP = Trim(GTsecuf)

End Function

Public Function GTTSK()

    sh = DatePart("h", Date)
    sm = DatePart("n", Date)
    ss = DatePart("s", Date)

XCodobJ = sh & sm & ss & Rnd()

GThms = DateDiff("d", Date, CDate("12/12/2050")) & FormatDateTime(Now(), 3)
GTses = Session.sessionid
GTsecu = CStr(XCodobJ & GThms & GTses)

GTsecuf = "TSK"

For GTi = 1 To Len(GTsecu)
GTcs = Mid(GTsecu, GTi, 1)
If GTcs >= "0" And GTcs <= "9" Then
   GTsecuf = GTsecuf & GTcs
End If
Next
 
GTTSK = Trim(GTsecuf)

End Function


end class
%>
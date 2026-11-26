
<%

 FUNCTION GENFMTWEDT(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""


'''''''''''''''''''''''''''''''''''''''''''''''''

  FNSHTML = FNSHTML & ofv.GenTabla("","","","100%","",0,0,0,0,"","","","no")


''''''''''''''''''
      stexto1 = "Capitulos"
      sOnClik = ""
      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
      Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")



  Strsql = "Select * From OBJNWEDT Where codobjeto=" & nCodigo & " Order By consecuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
   '  xxxx = NWEDTVER()
  end if

  XNVER = "N"
  XNSEC = 0


  do until rs.eof

  '''''FNSHTML = FNSHTML & rs(3)

     tobjetivo      = rs(3)

     if isnull(tobjetivo)   then tobjetivo    = " "

     tobjetivo      = "<BIG>" & replace(tobjetivo," ","&nbsp;") & "</BIG>"


     sObjetivo = ""


     sObjetivo = "" & getReadfile(rs(5) & ".CSE","")

  if sObjetivo <> "" then


'''''''''' arma tabla de indices '''''''''''''''''''''''''
      sPoint  = "T" & rs(0)
      stexto1 = rs(3)
      Celdat  = stexto1

      stexto1 = "<A class=ut href='javascript:void(0);' "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & sfidk1 & "</td>"
      Celdat  = Celdat & "<td class=uti onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
      Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



      pmk = "<form name="  & chr(34) & sPoint & chr(34) &  " method=post >"
      pmk = pmk & "<input type=hidden value=a ></form>"


        h1 = ofv.GenCelda("","","","300","","","","&nbsp;" & pmk,"si")
        FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
        h1 = ofv.GenCelda("","ttt  ","","","","","",tOBJETIVO,"si")
        FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
        sObjetivo=replace(sObjetivo,">&nbsp;</P>","><br></P>")
        h1 = ofv.GenCelda("","uti","","100%","","","",replace(sObjetivo,"&nbsp;"," ",1,-1,1),"si")
        FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


  end if



    rs.movenext
  loop

  FNSHTML = FNSHTML & ""
  FNSHTML = FNSHTML & ofv.GenfTabla()



  call ofv.cerrarconsulta(rs)


  GENFMTWEDT = "<div id=zglos class='norma-card'>" & FNSHTML & "</div>"


 END FUNCTION






function NWEDTVER()

  XNVER = "N"
  XNSEC = 0

  Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
  XferTarget0 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetasCMD\"
  XferTarget1 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "carpetaObjetivos\"
  XferTarget2 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaAlcance\"
  XferTarget3 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaConsideracion\"

  abrir1 = XferTarget1 & RtfObjetivo
  abrir2 = XferTarget2 & RtfAlcance
  abrir3 = XferTarget3 & RtfConsid

  set fff = server.createobject("scripting.filesystemobject")




'''''''''''''' Objetivo
''''
  Strsql = "Select ARCHIVO,titulo From ObjMain where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget1 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""

       xxx=getReadfile2(xarchivo ,"")
       set XNPSR = server.createobject("CorRtfParser.RtfParser")
       IF XXX <>"" THEN
        XNPSRFL   =  XNPSR.RtfToHtml(xxx)
       END IF
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

 '   end if
  end if
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Alcance
''''
  Strsql = "Select ARCHIVO,titulo From ObjAlcance where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget2 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""

      xxx=getReadfile2(xarchivo ,"")
     ' response.write xxx & "<br>" & xarchivo
      set XNPSR = server.createobject("CorRtfParser.RtfParser")
      IF XXX <> "" THEN
        XNPSRFL   =  XNPSR.RtfToHtml(xxx)
      END IF
      Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
      MyFile.Write XNPSRFL
      MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  '  end if
  end if
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Consideraciones generales
''''
  Strsql = "Select ARCHIVO,titulo From ObjConsid where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget3 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""

        xxx=getReadfile2(xarchivo,"") & ""
       ' response.write xxx & "<br>" & xarchivo
       set XNPSR = server.createobject("CorRtfParser.RtfParser")

       IF XXX <> "" THEN
        XNPSRFL   =  XNPSR.RtfToHtml(xxx)
       END IF
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  '  end if
  end if
  call ofv.cerrarconsulta(rs1)

  if  XNVER = "S" then

      Strsql = "Select * From OBJNWEDT Where codobjeto=" & nCodigo & " Order By consecuen"
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if

end function


%>
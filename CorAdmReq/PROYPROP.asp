<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0






  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar





  PRY = REQUEST.QUERYSTRING("PRY")

  IF PRY <> "" THEN  session("NCliente") = PRY

  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND TIPO = 'AN' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     GENSECTOROR()


  END IF


  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND TIPO = 'UC' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     GENUSRCONT()


  END IF



  XCONTROL = "N" 
  XCPT     = "N" 
  XCADV    = "N" 

  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' AND TIPO = 'AN' "
  set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if rsx1.eof then

     StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' "
     set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if



  if NOT rsX1.eof then

     IF RSX1(11) THEN
        XCONTROL = "S" 
        XCPT     = "S" 
        XCADV    = "S" 
     ELSE
        IF RSX1(13) THEN XCPT      = "S"
        IF RSX1(12) THEN XCADV     = "S"
     END IF

  END IF

  IF XCONTROL = "N" THEN

     bAgregar   = 0
     bModificar = 0
     bSuprimir  = 0

  END IF




 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

 FUNCION = ""




  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "CPRY" & chr(34) & "," & chr(34) & "resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "

  shtml = shtml & " var cnt = 0; " & chr(13)
  shtml = shtml & " function dos() {  " & chr(13)
  shtml = shtml & " if (document.all.odiv) { document.all.odiv.style.top=document.body.clientHeight - document.all.odiv.clientHeight; " & chr(13)
  shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
  shtml = shtml & " setTimeout (" & chr(34) & "dos()" & chr(34) & ", 350); }}" & chr(13)
  shtml = shtml & " } " & chr(13)
  shtml = shtml & " </script>" & chr(13)



    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenN = "<B><FONT COLOR=#993333 FACE=WEBDINGS SIZE=3>i</FONT></B>"
    sImagenR = "<B><FONT COLOR=#993333 FACE=WEBDINGS SIZE=3>&#0254;</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
    sImagenPR = "<B><FONT FACE=SYMBOL COLOR=#AA0000 SIZE=4>!</FONT></B>" 

    scomK = "0" & chr(9) & "1"
    scomD = "DIAS CORRIDOS" & chr(9) & "DIAS HABILES"


CARGAR = "S"

  IF CARGAR = "S" THEN

  StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 

     VPLAN = RS(17)
     SESSION("VPLAN")  = VPLAN
     SESSION("PRYCTA") = RS(20)



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas  = ofv.GenCelda("","ttt colspan=2","","","","","",sImagen & "&nbsp;PROPIEDADES DEL PROYECTO","si")
  filasm  = ofv.GenRow("","","","","10","","",celdas,"si")


'''''''''''''''''''''''''''''''''''''


    sImagen = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"

  celdas = ofv.GenCelda("","ut","","25%","","","","CODIGO:","si")
  celdas = celdas & ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")

  celdas = celdas & ofv.GenCelda("","ut","","70%","","","",rs(1),"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","ut","","","","","","TITULO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text","Usuarios","50","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","50","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","CLASE:","si")

    IF RS(15) = "REQ" THEN

       SINPUT1 = sImagenR
      sniv = "DESCONOCIDO"
      strsql = "Select * From REQUERIMIENTO where ID = " & rs(16)
      set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rsniv.eof then
         sniv = rsniv(19) & " /<U>" & rsniv(2) & "</U>"
         xlink = "<A CLASS=CC href='&nbsp;'  "
         xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsniv(1) & "&tipo=REQ'); return false;" & chr(34)
         xlink = xlink & " >" & sniv & "</A>"
         SNIV = XLINK
      END IF
      call ofv.cerrarconsulta(rsniv)



      SINPUT2 = "ASOCIADO A REQUERIMIENTO:&nbsp;" & SNIV
       

    ELSEIF RS(15) = "NO" THEN
       SINPUT1 = sImagenN    
       SINPUT2 = "PROYECTO INDEPENDIENTE"
    ELSEIF RS(15) = "OR" THEN
       SINPUT1 = sImagenR    
       SINPUT2 = "ASOCIADO A PLAN DE MITIGACION"
    ELSE
       SINPUT1 = "&nbsp;"     
       SINPUT2 = "ASOCIADO A DOCUMENTO DESCONOCIDO "
    END IF

    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput1,"si")

    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput2,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

''''''''''''
    if RS(9) = 0 then
      celdas = ofv.GenCelda("","ut","","10%","","","","TEMA/SUBTEMA:","si")
      celdas = celdas & ofv.GenCelda("","ut","","","","","","&nbsp;","si")


    if bModificar then

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYPROPTS.asp  "
     sinput = sinput  & "  TITLE='ASIGNAR'  ><b>" & sImagenI3 & "&nbsp;ASIGNAR</b></a>" 

    ELSE

     sinput = ""

    END IF


      celdas = celdas & ofv.GenCelda("","vt  ","","","","","","SIN DETERMINAR&nbsp;&nbsp;" & sinput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    else
      XTST = SPLIT(GETTEMASUBTEMA(RS(9)),CHR(10))
      SINPUT1 = XTST(0)
      SINPUT2 = XTST(1)
      

      celdas = ofv.GenCelda("","vt ","","","","","","TEMA:","si")
      celdas = celdas & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","vt  ","","","","","",SINPUT1,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt ","","","","","","SUBTEMA:","si")
      celdas = celdas & ofv.GenCelda("","ut","","","","","","&nbsp;","si")

    if bModificar then

      sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
      sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
      sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
      sinput = sinput  & "none" & chr(34) & ";' "
      sinput = sinput  & "  href=PROYPROPTS.asp  "
      sinput = sinput  & "  TITLE='MODIFICAR'  ><b>" & sImagenI3 & "&nbsp;MODIFICAR</b></a>" 

    ELSE

     sinput = ""

    END IF

      celdas = celdas & ofv.GenCelda("","vt  ","","","","","",SINPUT2 & "&nbsp;&nbsp;" & sinput,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


 
    end if
''''''''''''

    celdas = ofv.GenCelda("","ut","","","","","","PLAN VIGENTE:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

      strsql = "Select * From PROYPLAN where VPLAN = " & VPLAN & " AND CODIGO = '" & session("NCliente") & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(3)
        call ofv.cerrarconsulta(rsniv)
      sInput = VPLAN & ". " & sniv

     SESSION("VPLAND") = sInput 

    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    scomK = "0" & chr(9) & "1"
    scomD = "INCIDENCIAS" & chr(9) & "PROYECTOS/REQUERIMIENTOS"


    celdas = ofv.GenCelda("","ut","","","","","","MODO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then


        sInput = ofv.gencomboFX("Usuarios",rs(21).name,rs(21),scomK,scomD," ")

    else
        if rs(21) = 0 then sInput = "INCIDENCIAS"
        if rs(21) = 1 then sInput = "PROYECTOS/REQUERIMIENTOS"
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","ut","","","","","","CUENTA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
         StrSql = "Select * From PROYCUENTAS WHERE ESTADO = 1 ORDER BY ORDEN"
         sInput = ofv.GenerarCombo("Usuarios",rs(20).NAME,rs(20),strsql,cn,0,1,"")
      else
         StrSql = "Select DESCRIPCION From PROYCUENTAS where ID = " & RS(20)
         set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if NOT rsX2.eof then
            sInput  = rsX2(0)

         ELSE
            sInput  = "SIN ASIGNAR"

         END IF
      end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","TIPO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      strsql = "Select * From TIPOS Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(8).name,rs(8),strsql,cn,0,1,"")
    else
      strsql = "Select * From TIPOS where ID = " & rs(8)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","ORIGEN:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      strsql = "Select * From ORIGEN Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(7).name,rs(7),strsql,cn3,0,1,"")
    else
      strsql = "Select * From ORIGEN where OBJID = 0" & rs(7)
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","DESCRIPCION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(12).name,rs(12),"text","Usuarios","50","dt","")
    else
      sInput = ofv.GenerarInput("",rs(12),"readonly","","50","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    filam1 = filas
    tab1m = ofv.GenTabla("","","","100%","","0","","0","0","","",filam1,"si")




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    filas = ""

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


      celdas = ofv.GenCelda("","vt  ","","","","","","PLANIFICACION:","si")
      celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar AND RS(11) = 1 then
        sInput = "<font class=at><select name=" & rs(18).name & " onchange=submit()>"
        sInput = sInput & "<option value=0 "
        if rs(18) = 0 then sInput = sInput & "selected"
        sInput = sInput & " >DIAS CORRIDOS<option value=1 "
        if rs(18) = 1 then sInput = sInput & "selected"
        sInput = sInput & " >DIAS HABILES</select></font>"

        sInput = ofv.gencomboFX("Usuarios",rs(18).name,rs(18),scomK,scomD," ")
    else
        if rs(18) = 0 then sInput = "DIAS CORRIDOS"
        if rs(18) = 1 then sInput = "DIAS HABILES"
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","PRIORIDAD:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagenPR,"si")




    if bModificar then
       StrSql = "Select * From PRIORIDAD ORDER BY ID"
      sInput = ofv.GenerarCombo("Usuarios",rs(19).name,rs(19),strsql,cn,0,1,"")
    else
     StrSql = "Select * From PRIORIDAD where ID = '" & RS(19) & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        sInput  = rsX2(1)

     ELSE
        sInput  = "SIN DETERMINAR"

     END IF
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","CREACION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","INICIO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(4)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(4) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    StrSql = "Select tarea From PROYETPTSK where codigo = '" & session("NCliente") & "' and VPLAN = " & session("vplan") & " ORDER BY ORDEN desc,TSKORD DESC "
    set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT rs5.EOF THEN
   '    OWKF.YFMAXTSK = rs5(0)

'       SXFSTARTD = OWKF.GETFENDV(rs5(0), session("NCliente"))
'
 '      XIIX = OWKF.TIX


'       FSPLAN = OWKF.ATABTSK(cint(XIIX),7)
    ''   FEPLAN = OWKF.ATABTSK(cint(XIIX),8)
 '      FSREAL = OWKF.ATABTSK(cint(XIIX),9)

'       FEPLAN = rs(5)
 '      FEREAL = OWKF.ZXMFMAX
       
    else
  '     SXFSTARTD = OWKF.GETFENDV(session("NCliente"), session("NCliente"))

'       XIIX = OWKF.TIX


'       FSPLAN = OWKF.ATABTSK(cint(XIIX),7)
''       FEPLAN = OWKF.ATABTSK(cint(XIIX),8)
'       FSREAL = OWKF.ATABTSK(cint(XIIX),9)

       FEPLAN = rs(5)
'       FEREAL = OWKF.ATABTSK(cint(XIIX),10)

    end if


    call ofv.cerrarconsulta(rs5)




    celdas = ofv.GenCelda("","ut","","","","","","ENTREGA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(FEPLAN) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(FEPLAN + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","FINALIZACION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(FEREAL) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(FEREAL + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","ESTADO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","<B>" & sImagenE & "</B>","si")

      sInstrumento = "---"

        Strsql = "Select descri From reqestados Where id = " & rs(11)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sInstrumento


    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","ut","","","","","","ORIGINANTE:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","<B>" & sImagenO & "</B>","si")
    if RS(13) = "NO" then
      sInput = "SIN ORIGINANTE"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","RESPONSABLE:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","<B>" & sImagenI & "</B>","si")
    if RS(14) = "NO" then
      sInput = "SIN ASIGNAR"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & rs(14) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = "<B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    filam2 = filas
    tab2m = ofv.GenTabla("","","","100%","","0","","0","0","","",filam2,"si")




''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      celdasm = ""

      celdasm = celdasm & ofv.GenCelda("","ut VALIGN=TOP style='border-style:solid;border-width:1;border-color:#dddddd;' ","","50%","","","",tab1m,"si")
      celdasm = celdasm & ofv.GenCelda("","ut VALIGN=TOP style='border-style:solid;border-width:1;border-color:#dddddd;' ","","50%","","","",tab2m,"si")
      filasm  = filasm & ofv.GenRow("","","","","","","",celdasm,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYECTO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filasm,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filasm,"si")
    end if

    formulario0 = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;' ","","100%","","0","","2","0","","",formulario0,"si")


    sHTML = sHTML  & formulario0 

    sHTML = sHTML & "</center>"



  IF XCONTROL <> "N" THEN

       celdas = ""


       if rs(11) = 1 then

          sInput  = ofv.GenerarInput("boton","INICIAR","submit","","10","bt","")
          sAccion = "proyini.asp"
          btcel   = ofv.GenCelda("","UT TITLE='Iniciar proyecto.' ","","","","","",sInput,"si")
          celdas  = celdas & ofv.GenForm("init","",sAccion,"","post","",btcel,"si")

       end if


       if rs(11) = 12 then
          StrSql = "Select * From PROYETPTSK where codigo = '" & session("NCliente") & "' and ESTADO < 2 AND VPLAN = " & VPLAN 
          set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          IF rs5.EOF THEN
             sInput  = ofv.GenerarInput("boton","FINALIZAR","submit","","10","bt","")
             sAccion = "proyEND.asp"
             btcel   = ofv.GenCelda("","UT TITLE='FINALIZAR EL PROYECTO.' ","","","","","",sInput,"si")
             celdas  = celdas & ofv.GenForm("init2","",sAccion,"","post","",btcel,"si")
          END IF

       end if



    celdas = celdas & ofv.GenCelda("","ut","center","","90%","","","&nbsp;","si")


    filas = ofv.GenRow("","","","","10","","",celdas,"si")


    sHTML = sHTML  & "<br>" & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;' ","","100%","","0","","2","0","","",filas,"si")


  END IF





 
  end if

 END IF

  'Botonera
  Agregar = ""
  Buscar  = ""
    Celda   = ""

   '' Celda   = ofv.BotonVolver(sAccion)     


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GETTEMASUBTEMA(SSUBTEMA)

  xTEMA = 0
  xSUBTEMA = "SIN ASIGNAR"

  SQL2 = "SELECT * "
  SQL2 = SQL2 & " FROM csbuscasubtema where "
  SQL2 = SQL2 & " codPRODSERV = " & SSUBTEMA
  set rsX = ofv.crearconsultaEx(SQL2,cn3,1,parametros)
  if NOT rsX.eof then
     xTEMA = RSX(3)
     xSUBTEMA = RSX(1)
     SQL1 = "SELECT * From Tema WHERE CODTEMA = " & xTEMA
     set rsX = ofv.crearconsultaEx(SQL1,cn3,1,parametros)
     if NOT rsX.eof then
        xTEMA = RSX(1)
     END IF

  END IF

  GETTEMASUBTEMA = XTEMA & CHR(10) & XSUBTEMA

END FUNCTION



%>
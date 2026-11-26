<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.MenuHeader("",funcion)
 

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

'  IF session("ModMain") <> "WF" THEN 
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

 XMOD = request.querystring("XMOD")
 XDST = request.querystring("XDST")

 XSET = request.querystring("XSET")
 XOPT = request.querystring("XOPT")
 sopcion = request.querystring("opcion") 

 sRESP = request.querystring("RESP")

    DIM ATAREA(200,30)
    DIM ATAREA2(200,30)
    DIM ATAREA3(200,30)

   ITART = 0
   ITATI = 0
   TTAR  = 0

   IETPT = 0
   IETPI = 0 
   TETP  = 0

   CDESP = 0
   CHASP = 0
   CDESR = 0
   CHASR = 0
   CMAX  = 0  




     sopcion = "VP"


    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4s = "<FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>&#0128;</FONT>"
    sImagen4n = "<FONT FACE=WEBDINGS COLOR=#FFFFFF SIZE=2>x</FONT>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagenXLS = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#048;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>&#0233;</FONT></B>"
    sImagen6A = "<B><FONT FACE=WINGDINGS COLOR=#F7AA22 SIZE=3>&#0190;</FONT></B>"
    sImagen6W = "<B><FONT FACE=WINGDINGS COLOR=#CCCCCC SIZE=2>s</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"
    sImagenE4 = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=4>&#0232;</FONT></B>"
    sImagenFF = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=4>4&#0166;</FONT></B>"

    sImageG = "<B><FONT FACE=WEBDINGS COLOR=#6699CC SIZE=3>=</FONT></B>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"
    sImageR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0128;</FONT></B>"



SELECT CASE XOPT


 CASE "XGLB"

       XDEST="GLOBAL"

  Strsql = "Select * from clienteplan where id = 1"
  set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
  if not rs.eof then
     NOMEMP = rs(1)
     LOGEMP = rs(5)
  end if
  call ofv.cerrarconsulta(rs)

       XDESTN=NOMEMP


  CASE "XGRN"

       XDEST="GERENCIA"

     StrSql = "Select * From ECO where ECO = " & SRESP
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XDESTN=sniv

  CASE "XSCT"

       XDEST="SECTOR"

     StrSql = "Select * From ECO where ECO = " & SRESP
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XDESTN=sniv

  CASE "XRESP"

       XDEST="RESPONSABLE"

      strsql = "Select * From USUARIOS where ID = '" & SRESP & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then
         sniv = rsniv(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsniv)


       XDESTN=sniv

END SELECT

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSFEC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & "   TITLE='MODIFICAR RANGO DE SELECCION'  ><b>&nbsp;" & sImagenFF & "</b></a>" 


if session("FECOPC") = "1" THEN
   XSELN = "&nbsp;&nbsp;&nbsp;&nbsp;SELECCION:&nbsp;<B>TODOS</B>&nbsp;" & sinput

   SESSION("FECDES") = 0
   SESSION("FECHAS") = 99999 

ELSE


       dFechaD = UCASE(Fechador(SESSION("FECDES")))
       dFechaH = UCASE(Fechador(SESSION("FECHAS")))

 '        dFechaD = CDATE(CLNG(SESSION("FECDES")) + clng(cdate("01-01-1900")))
 '        dFechaH = CDATE(CLNG(SESSION("FECHAS")) + clng(cdate("01-01-1900")))

   XSELN = "&nbsp;&nbsp;&nbsp;&nbsp;SELECCION:&nbsp;<B>RANGO</B><BR>&nbsp;DESDE&nbsp;<B>" & dFechaD & "</B>&nbsp;&nbsp;&nbsp;HASTA&nbsp;<B>" & dFechaH & "</B>&nbsp;" & sinput

END IF

 'response.write "recibe " & request.querystring & "<br>"


  ITART   = 1
  ITART2  = 1
  ITART3  = 0
  TTAR    = 1



        ATAREA (ITART,2)    = "H"
        ATAREA (ITART,3)    = 1
        ATAREA (ITART,4)    = 1

        ATAREA2 (ITART2,0)  = 0
        ATAREA2 (ITART2,1)  = "NIVEL:&nbsp;<B>" & XDEST & "</B>&nbsp;&nbsp;&nbsp;&nbsp;DESTINO:&nbsp;<B>" & XDESTN & "</B>" & XSELN
        ATAREA2 (ITART2,2)  = "S"
        ATAREA2 (ITART2,3)  = 1
        ATAREA2 (ITART2,4)  = 0
        ATAREA2 (ITART2,5)  = sImageG





  SELECT CASE XMOD


        CASE "TAR"






        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "ASIGNACION DE PROYECTOS/TAREAS"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0
        
        
     sinputa = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinputa = sinputa  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinputa = sinputa  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinputa = sinputa  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinputa = sinputa  & "none" & chr(34) & ";' "
     sinputa = sinputa  & "  href=PROYASGN.asp?XLS=S&XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinputa = sinputa  & " TARGET=_blank  TITLE='EXPORTAR A EXCEL" & "" & "'  ><b>" & sImagenXLS & "</b></a>" 
     

     sinput = sinputa & "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYASGN.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='ASIGNACION DE PROYECTOS/TAREAS DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
     ATAREA3 (ITART3,5)  = sinput


        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "GANTT DE TAREAS"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYASGNGRAF.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='GANTT DE TAREAS DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "TOTAL DE HORAS POR RESPONSABLE"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSTHR.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='TOTAL DE HORAS POR RESPONSABLE DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "TOTAL DE HORAS INSUMIDAS EN EL DIA"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSTHC.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='TOTAL DE HORAS INSUMIDAS EN EL DIA DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

     
     ITART3             = ITART3 + 1
     ATAREA3 (ITART3,0)  = 0
     ATAREA3 (ITART3,1)  = "TOTAL DE HORAS INSUMIDAS EN EL DIA NUEVO"
     ATAREA3 (ITART3,2)  = "N"
     ATAREA3 (ITART3,3)  = 0
     ATAREA3 (ITART3,4)  = 0
        
        
     sinputa = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinputa = sinputa  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinputa = sinputa  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinputa = sinputa  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinputa = sinputa  & "none" & chr(34) & ";' "
     sinputa = sinputa  & "  href=PROYCNSTHCNew.asp?XLS=S&XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinputa = sinputa  & " TARGET=_blank  TITLE='EXPORTAR A EXCEL" & "" & "'  ><b>" & sImagenXLS & "</b></a>" 
     
     sinput = sinputa & "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSTHCNew.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='TOTAL DE HORAS INSUMIDAS EN EL DIA DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
     ATAREA3 (ITART3,5)  = sinput

        
        
        ITART3             = ITART3 + 1
        
        
        
        
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "DISTRIBUCION DE HORAS POR TIPO DE TAREA"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSDHT.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='DISTRIBUCION DE HORAS POR TIPO DE TAREA DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "GANTT DE PROYECTOS"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYPRYGRAF.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='GANTT DE PROYECTOS DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "GAP DE ASIGNACION DE RECURSOS"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYASGNGAP.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank TITLE='GAP DE ASIGNACION DE RECURSOS DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

     IF XOPT = "XRESP" THEN

        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "LISTA DE ASIGNACION"
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYASGNLIST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank TITLE='LISTA DE ASIGNACION DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

    END IF

     CASE ELSE


        IF XMOD = "PRY" THEN
           XDSTDSC = "PROYECTOS"
        ELSE
           XDSTDSC = "REQUERIMIENTOS"
        END IF




        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "DETALLE GENERAL DE " & XDSTDSC
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYPRYREQLST.asp?XSET=" & XSET & "&XDET=1&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='DETALLE GENERAL DE " & XDSTDSC & " DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput


        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "TOTAL GENERAL DE " & XDSTDSC
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYPRYREQLST.asp?XSET=" & XSET & "&XDET=0&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='TOTAL GENERAL DE " & XDSTDSC & " DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

    IF XMOD = "REQ" THEN



        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "DETALLE ANALITICO DE " & XDSTDSC
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYPRYREQLSTAN.asp?XSET=" & XSET & "&XDET=1&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='DETALLE ANALITICO DE " & XDSTDSC & " DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput


        ITART3             = ITART3 + 1
        ATAREA3 (ITART3,0)  = 0
        ATAREA3 (ITART3,1)  = "TOTAL ANALITICO DE " & XDSTDSC
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYPRYREQLSTAN.asp?XSET=" & XSET & "&XDET=0&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=" & XOPT & "&RESP=" & SRESP
     sinput = sinput  & " TARGET=_blank  TITLE='TOTAL ANALITICO DE " & XDSTDSC & " DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

    END IF


  END SELECT


    ATAREA2 (ITART2,4)  = ITART3






    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 


      celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>REPORTES " & SESSION("FORMN") & "&nbsp;&nbsp;&nbsp;&nbsp; AL " & ofv.fechador & "</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")





      celdas = ofv.GenCelda("","VT COLSPAN=4 STYLE='border-bottom-style:RIDGE;border-bottom-width:2;border-bottom-color:#f7aa22;' ","CENTER","90%","","","","RESPONSABLES","si")

      filas = ofv.GenRow("","","","","10","","",celdas,"si")


''''''''''''''''


  for ITART=1 to ttar

      IF ATAREA (ITART,2) <> "H"  THEN

      celdas = ofv.GenCelda("","VT ","right VALIGN=TOP ","","","","",ATAREA (ITART,5),"si")
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=3  VALIGN=TOP ","","","","","",ATAREA (ITART,1),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
    
      END IF

      IF ATAREA (ITART,2) = "S" OR ATAREA (ITART,2) = "H" THEN

         for ITART2=ATAREA (ITART,3) to ATAREA (ITART,4)

             IF ATAREA2 (ITART2,2) <> "H"  THEN

             celdas = ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
             celdas = celdas & ofv.GenCelda("","","right VALIGN=TOP ","","","","",ATAREA2 (ITART2,5),"si")
             celdas = celdas & ofv.GenCelda("","ut COLSPAN=2  VALIGN=TOP ","","","","","",ATAREA2 (ITART2,1),"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

             END IF

             IF ATAREA2 (ITART2,2) = "S" OR ATAREA2 (ITART2,2) = "H"  THEN

                for ITART3=ATAREA2 (ITART2,3) to ATAREA2 (ITART2,4)

                    celdas = ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
                    celdas = celdas & ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
                    celdas = celdas & ofv.GenCelda("","","right VALIGN=TOP  ","","","","",ATAREA3 (ITART3,5),"si")
                    celdas = celdas & ofv.GenCelda("","ut  VALIGN=TOP  ","","","","","",ATAREA3 (ITART3,1),"si")
                    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

               next

             END IF


          next

      END IF


  next



      celdas = ofv.GenCelda("","","CENTER","","2%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","10%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","10%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","80%","","","&nbsp;","si")

    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    FilaS2 = ofv.GenTabla("","aa style='border-style:ridge;border-width:2;border-color:#f7aa22;' ","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","AAA COLSPAN=2","center","","","","",FilaS1 & FilaS2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

     cnbot = "<input type=submit class=bt value='VOLVER' >"
     Cons = ofv.gencelda("","","","","1%","","",cnbot,"si")
     sAccion = "PROYCNSTREE.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  
     Cons = ofv.genform("cons","",sAccion,"","post","",Cons,"si")
     celdas = Cons

      celdas = celdas & ofv.GenCelda("","","CENTER","","80%%","","","&nbsp;","si")

    FilaS = FilaS & ofv.GenRow("","","","","","","",CeldaS,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





'''''''''''''''












 







  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>
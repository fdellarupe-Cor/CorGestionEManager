<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = "../WKFPASOSPROPNV.ASP"

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")
 WKFPSCTPC = request.querystring("WKFPSCTPC")

 if WKFPSCTPC = "" then
    WKFPSCTPC = session("WKFPSCTPC")
 else
    session("WKFPSCTPC") = WKFPSCTPC
 end if 

    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenF = "<B><FONT FACE=webDINGS COLOR=#993333 SIZE=5>q</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
    sImagenSI = "<B><FONT FACE=WINGDINGS COLOR=#0000aa SIZE=4>&#0236;</FONT></B>"
    sImagenSF = "<B><FONT FACE=WINGDINGS COLOR=#000000 SIZE=4>&#0237;</FONT></B>"
    sImagenSA = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=5>&#0252;</FONT></B>"
    sImagenSR = "<B><FONT FACE=WINGDINGS COLOR=#aa0000 SIZE=5>&#0251;</FONT></B>"
    sImagenPA = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=4>&#0231;</FONT></B>"
    sImagenPS = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=4>&#0232;</FONT></B>"


IF SOPCION = "SI" OR SOPCION = "CM" THEN

   IF SOPCION = "SI"  THEN

      NEWPASO = GENWKFPASONEW()

   ELSE

      NEWPASO = request.querystring("WKFNEWPASO")

   END IF





   STRSQL = "UPDATE WKFPASOSPRTYNV SET PASOSIG = '" & NEWPASO & "' "
   STRSQL = STRSQL & " WHERE ID = " & session("WKFPSCTPC")
   call ofv.crearconsultaEx(StrSql,cn,1,volver)

   RESPONSE.REDIRECT "WKFPASOSPROPNV.ASP"

ELSE


  StrSql = "Select * From WKFCIRCUITO WHERE WKFCOD = '" & session("WKF") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then



    sInput = "CIRCUITO:" & RS(1) & " - " & RS(2)
    celdas = ofv.GenCelda("","ttT","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","22","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


  END IF



    celdas = ""





  StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then



    sInput = "PASO:"
    celdas = ofv.GenCelda("","Ut","","10%","","","",sInput,"si")

    sInput = "<big>" & RS(3) & ". " & RS(4) & "</big>"
    celdas = celdas & ofv.GenCelda("","Ut","","","","","",sInput,"si")



    filas = ofv.GenRow("","","","","","","",celdas,"si")




  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


 END IF




  StrSql = "Select * From WKFPASOSPRTYNV WHERE ID = " & session("WKFPSCTPC")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then





     IF RS(3) = "INI" THEN sInput = "INICIAR"     
     IF RS(3) = "APR" THEN sInput = "APROBAR"     
     IF RS(3) = "RCH" THEN sInput = "RECHAZAR"     

    sInput = "DERIVAR&nbsp;POR&nbsp;OPCION:&nbsp;" & SINPUT & "&nbsp;A:&nbsp;&nbsp;"
    celdas = ofv.GenCelda("","Ut","","15%","","","",sInput,"si")


   WKPASO = RS(7) & ""
   if WKPASO <> "" then 
     sInput = GENWKFPASO(WKPASO)
   ELSE
    sInput = "SELECCIONAR"
   END IF

    sInput = sImagenPS & "&nbsp;" & SINPUT	

    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"





''' ANTERIOR

        sinputK = "<a class=nt style='text-decoration:none;color:#993333;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='WKFPSCSELPASNV.asp?OPCION=SI' "
        sinputK = sinputK  & "  TITLE='Crear nuevo paso'  ><b>" & sImagenPS & "&nbsp;&nbsp;<big>" & sImagenPC & "&nbsp;CREAR NUEVO PASO</big></b></a>" 

        STABC = ofv.GenCelda("","ut  ","","","","","",sInputK,"si")
        STAB = ofv.GenRow("","","","","","","",STABC,"si")
      strsql = "Select * From WKFPASOSNV WHERE WKFCOD = '" & session("WKF") & "' AND WKFPASCOD <> '" & RS(2) & "' AND WKFPASCOD <> '" & RS(7) & "'  Order By ORDEN, DESCRIPCION "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF not RSNIV.EOF THEN

      DO WHILE NOT RSNIV.EOF


          xpaso = GENWKFPASO(RSNIV(2))

        sinputK = "<a class=nt style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='WKFPSCSELPASNV.asp?OPCION=CM&WKFNEWPASO=" & RSNIV(2) & "' "
        sinputK = sinputK  & "  TITLE='Derivar a " & xpaso & "'  ><b>" & sImagenPS & "&nbsp;&nbsp;<big>" & xpaso & "</big></b></a>" 


        
        STABC = ofv.GenCelda("","ut  ","","","","","",sInputK,"si")
        STAB = STAB & ofv.GenRow("","","","","","","",STABC,"si")

         RSNIV.MOVENEXT
      LOOP
        STAB = ofv.GenTabla("","","","80%","","0","","0","0","","",STAB,"si")
        END IF
      call ofv.cerrarconsulta(rsniv)


''' ANTERIOR

    filas = ""

    celdas = ofv.GenCelda("","vt STYLE='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-color:#DDDDDD;' ","CENTER","15%","","","","<big>SELECCIONAR</big>","si")
    celdas = celdas & ofv.GenCelda("","ut  ","CENTER","85%","","","",STAB,"si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")
    STAB = ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

    celdas = ofv.GenCelda("","","CENTER","","","","",STAB,"si")
    filas = ofv.GenRow("","","","","","","",celdas,"si")
 
    sHTML = sHTML  & "<CENTER>"


  sHTML = sHTML  & ofv.GenTabla("","AAA STYLE='BORDER-STYLE:solid;BORDER-WIDTH:1;BORDER-color:#888888;' ","","100%","","0","","0","0","","",filas,"si") & "<BR>"



END IF





'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




  Agregar = ""
  Buscar  = ""

  sAccion = "WKFPASOSPROPNV.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",Agregar,Buscar,celda)




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"

 END IF

end if

response.write sHTML

FUNCTION GENWKFPASONEW()

    WKFCOD = GTWKFPSC()

  StrSql = "Select MAX(ORDEN) From WKFPASOSNV WHERE WKFCOD = '" & session("WKF") & "' and orden < 999 "
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then 
     NORDEN = RS3(0) + 10
  ELSE

     NORDEN = 20
  END IF

    strsql = "insert into WKFPASOSNV ("
    strsql = strsql & "WKFCOD,WKFPASCOD,ORDEN,DESCRIPCION,"
    strsql = strsql & "VALIDO,INI,FIN,SALTEAR,"
    strsql = strsql & "CANCELAR,NODO01,NODO02,NODO03,"
    strsql = strsql & "NODO04,NODO05,NODO06,NODO07,"
    strsql = strsql & "NODO08,NODO09,NODO10,TIPOCICLO,"
    strsql = strsql & "NIVELRSP,NROCICLO,TIEMPOESTIMADO,UNIDADTIEMPO,WKFACCION"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & session("WKF") & "','" & WKFCOD & "','" & NORDEN & "','" & NORDEN & " - NUEVO PASO WKF',"
    strsql = strsql & "'S','N','N','N',"
    strsql = strsql & "'N','S','N','N',"
    strsql = strsql & "'N','N','N','N',"
    strsql = strsql & "'N','N','N','U',"
    strsql = strsql & "'0','1','1','1','D'"
    strsql = strsql & ")"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    xxxx = GENWKFPSCTPC("3","N01")

  GENWKFPASONEW = WKFCOD

END FUNCTION

FUNCTION GENWKFPASO(XPASO)

  GENWKFPASO = "SIN ASOCIAR"

  StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & XPASO & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     sinput = RS2(3) & ". " & RS2(4)
     GENWKFPASO = sinput
  END IF
  call ofv.cerrarconsulta(rs2)

END FUNCTION

FUNCTION GENWKFPSCTPC(XTIPO,XITEMCOD)


    WKFCODPRTY = GTWKFPSC() & ITEMCOD

    IF XTIPO = "" THEN XTIPO = "3"
    IF XITEMCOD = "" THEN XITEMCOD = "N01"

    StrSql = "Select * From WKFPSCTIPOS WHERE TIPO = '" & XTIPO & "' "
    set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs3.eof then 
       XTIPODESC = RS3(2)
    ELSE

       XTIPODESC = "NO ENCONTRADO"
    END IF


    strsql = "insert into WKFPASOSPRTYNV ("
    strsql = strsql & "WKFCOD,WKFPASCOD,WKFPASPRTYCOD,DESCRIPCION,"
    strsql = strsql & "TIPOPSC,ITEMCOD,PASOSIG,PERMITECLC,"
    strsql = strsql & "CAMEST,ESTADO,WKFACCIONAD"
    strsql = strsql & ") values ("
    strsql = strsql & "'" & session("WKF") & "','" & session("WKFPSC") & "','" & WKFCODPRTY & "','" & XTIPODESC & "',"
    strsql = strsql & "'" & XTIPO & "','" & XITEMCOD & "','','S',"
    strsql = strsql & "'N','0','D'"
    strsql = strsql & ")"
 '   response.write strsql & "<br>" & chr(13)
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)





END FUNCTION


%>
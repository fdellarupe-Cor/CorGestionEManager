<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(session("FormN"))


    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=4>&#0154;</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"

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

 sopcion = request.querystring("opcion")

 if sopcion = "SI" THEN
    USR = request.querystring("USR")
     GENCONTACTO(USR)
 END IF




  StrSql = "Select * From REQCONTACTOS where codigo = '" & session("NCliente") & "' ORDER BY ID "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)





    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;CONTACTOS","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","TTR  ","center","","","","","","si")
  celdas = celdas & ofv.GenCelda("","TTR  ","center","","","","","CONTACTO","si")
  celdas = celdas & ofv.GenCelda("","TTR  ","center","","","","","SECTOR","si")
  filas = filas & ofv.GenRow("","","","","22","","",celdas,"si")



  DO UNTIL RS.EOF

    'Eliminar
    if bSuprimir  then

          SELIM = "S" 
    else
          SELIM = "N" 
    end if


       IF SELIM = "S" THEN  
         sHref = "asp/EliminarRegistro.asp?ID=" & rs(0)
         sHref = sHref & "&Tabla=REQCONTACTOS&volver=" & vuelta & "&Eliminar=NSNC"
         CeldaS = ofv.BotonEliminar(" CONTACTO ",sHref)
         CeldaS = REPLACE(CeldaS,"<TD","<TD WIDTH=1%  ")
       ELSE
         CeldaS = ofv.GenCelda("","UT  ","","1%","","","","&nbsp;","si")
       END IF



         sniv = "INEXISTENTE"
         snivS = 0

      strsql = "Select * From USUARIOS where ID = '" & rs(2) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then
         sniv = rsniv(1)
         snivS = rsniv(21)
      END IF
      call ofv.cerrarconsulta(rsniv)
      sInput = sImagenI2 & "&nbsp;" & sniv & ""
     celdas = CELDAS & ofv.GenCelda("","UT","","","","","",sInput,"si")


     IF SNIVS > 0 THEN
        strsql = "Select * From ECO where ECO = " & SNIVS
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then
           sniv = rsniv(1)
        END IF
        call ofv.cerrarconsulta(rsniv)
        sInput = sniv
        celdas = CELDAS & ofv.GenCelda("","UT","","","","","",sInput,"si")  
     ELSE
        celdas = CELDAS & ofv.GenCelda("","UT","","","","","","SIN AFECTAR","si")
     END IF 

     filas = filas & ofv.GenRow("","","","","20","","",celdas,"si")

    RS.MOVENEXT
    LOOP


  sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","-1","-1","","",filas,"si")
  sHTML = sHTML & "</center><br><br>"


'  cnbot = "<input type=submit class=bt value='AGREGAR&nbsp;CONTACTO' >"
'  Cons = ofv.gencelda("","","center","","1%","","",cnbot,"si")
'  sAccion = "EDCONTACTOSBUS.asp?VASP=ES"
'  Buscar = ofv.genform("cons","",sAccion,"","post","",Cons,"si")

  'Botonera
   Buscar = ""
  if bAgregar then
     sAccion = "EDCONTACTOSBUS.asp?VASP=ES"
     Buscar = ofv.BotonAgregar(sAccion)
  end if



''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1
 ' Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



FUNCTION GENCONTACTO(XUSR)

        STRSQL = "INSERT INTO REQCONTACTOS (CODIGO,CONTACTO) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & XUSR & "')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
''        response.write strsql


END FUNCTION

%>
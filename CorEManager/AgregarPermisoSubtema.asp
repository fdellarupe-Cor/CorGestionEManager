<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Volver = "../subtemaskr.asp"
  sxopc1 = ""


  SCOMSEP = ""
  scomK   = ""
  scomD   = ""

'  set cn = ofv.conectar(ofv.strconn0)
'  strsql = "Select * From Grupos order by descripcion"
'  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
'  do until rs.eof
'   ''  sxopc1 = sxopc1 & "<option value=GRP-" & rs(0) & ">Grupo:" & rs(1) & "</option> "
'         if sInputAi = "" then sInputAi = "GRP-"   & rs(0)
'         scomK = scomK & SCOMSEP & "GRP-"   & rs(0)
'         scomD = scomD & SCOMSEP & "Grupo:" & rs(1)
'         SCOMSEP = chr(9)
'  rs.movenext
'  loop
'  strsql = "Select * From Usuarios  order by descripcion"
'  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
'  do until rs.eof
'   ''  sxopc1 = sxopc1 & "<option value=USR-" & rs(0) & ">Usuario:" & rs(1) & "</option> "
'         if sInputAi = "" then sInputAi = "USR-"   & rs(0)
'         scomK = scomK & SCOMSEP & "USR-"     & rs(0)
'         scomD = scomD & SCOMSEP & "Usuario:" & rs(1)
'         SCOMSEP = chr(9)
'  rs.movenext
'  loop
'  call ofv.cerrarconsulta(rs)
'  call ofv.cerrarconn(cn)



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","N","S","S",CHR(9),CHR(9),"GRP-","USR-")

     IF SESSION("FBUSR") <> "NO" THEN

        sxopc2 = ""

        set cn = ofv.conectar(ofv.strconn2)

        strsql = "Select * From prodserv where codprodserv = " & session("subt")
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
        if not rs.eof then

           sInputBi = "A"
           dsubt    = rs(1)

           SCOMSEP  = ""
           scomK2   = scomK2 & SCOMSEP & "A"
           scomD2   = scomD2 & SCOMSEP & "Acceso"
           SCOMSEP  = chr(9)

           if rs("segaceso") <> "0" or rs("segactual") <> "0" then
              scomK2 = scomK2 & SCOMSEP & "T"
              scomD2 = scomD2 & SCOMSEP & "Acceso y Actualizacion"

           end if

       end if
       call ofv.cerrarconsulta(rs)


       strsql = "Select * From subtemaskr "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)



       f1a = "asp/Agregar.asp?Tabla=subtemaskr"
       f1a = f1a & "&volver=" & volver

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

        sHTML = sHTML & "<center>"
        sHTML = sHTML & ofv.GenTabla("panel","aaa style='border-style:solid;border-width:1;border-color:#dddddd;background-color:#f0f0f0;' ","","650","",0,0,-2,-1,"","","","no")

        sInputA = "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(4).NAME & "' >"
        sInputB = ofv.gencomboFX(sForm,rs(3).name,sInputBi,scomK2,scomD2,"no")
        sInputC = ofv.GenerarInput(rs(1).name,session("subt"),"hidden","","","","")

       'Encabezados Descripcion y Secuencia
        h1 = ofv.GenCelda("","tt","","20%","","","","Permiso","si")
        h1 = h1 & ofv.GenCelda("","ut","","80%","","","",sInputA & sInputB & sInputC,"si")

        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
        sHTML = sHTML & ofv.GenFTabla()
        sHTML = sHTML & "</center>"
        call ofv.cerrarconsulta(rs)
        call ofv.cerrarconn(cn)


      ELSE

        sAccion = "subtemaskr.asp"

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botones Aceptar y Cancelar
  sAccion = "subtemaskr.asp"
  sHTML = sHTML & replace(ofv.AgregarBoton("AgreArchb",sAccion,"_self"),"<br><br><br>","")


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>
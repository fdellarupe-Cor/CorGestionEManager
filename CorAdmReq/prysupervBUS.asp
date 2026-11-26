<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Volver = "PRYSUPERV.asp"
  sxopc1 = ""


  SCOMSEP = ""
  scomK   = ""
  scomD   = ""
 




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

     SHTML = SHTML & GenBuscaUsr("N","N","S","S",CHR(9),CHR(9),"","")

     IF SESSION("FBUSR") <> "NO" THEN

        dsubt    = "0"
        dNOMB    = SESSION("FBUSR")

        set cn = ofv.conectar(ofv.strconn0)

        strsql = "Select * From USUARIOS where ID = '" & SESSION("FBUSR") & "' "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
        if not rs.eof then

           dsubt    = rs(21)
           dNOMB    = rs(1)

       end if
       call ofv.cerrarconsulta(rs)
       call ofv.cerrarconn(cn)




       f1a = "PRYSUPERV.asp?opcion=SI"

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

        sHTML = sHTML & "<center>"
        sHTML = sHTML & ofv.GenTabla("panel","aaa style='border-style:solid;border-width:1;border-color:#dddddd;background-color:#f0f0f0;' ","","650","",0,0,-2,-1,"","","","no")

        sInputA = ofv.GenerarInput("USRAFC",SESSION("FBUSR"),"hidden","","","","")
        sInputB = ofv.GenerarInput("SECTOR",dsubt,"hidden","","","","")


       'Encabezados Descripcion y Secuencia
        h1 = ofv.GenCelda("","tt","","20%","","","","AFECTAR","si")
        h1 = h1 & ofv.GenCelda("","ut","","80%","","","",SESSION("FBUSR") & " - " & dNOMB & sInputA & sInputB,"si")

        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
        sHTML = sHTML & ofv.GenFTabla()
        sHTML = sHTML & "</center>"

        call ofv.cerrarconn(cn)


      ELSE

        sAccion = "PRYSUPERV.asp"

        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")
  

      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botones Aceptar y Cancelar
  sAccion = "PRYSUPERV.asp"
  sHTML = sHTML & replace(ofv.AgregarBoton("AgreArchb",sAccion,"_self"),"<br><br><br>","")


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>
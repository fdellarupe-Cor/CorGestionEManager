<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

''''''''''''
  Generar = ""
  




  sHTML = "<html><body onload='window.top.ldfobj();' >"

  set cn = ofv.conectar(ofv.strconn2)

  sCODPROCESO = request.querystring("CODPROCESO") & ""
  sTIPPROCESO = request.querystring("TIPPROCESO") & ""

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_CODPROCESO=" & sCODPROCESO & "_TIPPROCESO=" & sTIPPROCESO


''''''''''''''''''''''''''''''''''''''''''''
  sForm = "EDPROCF"

  sImagenI3 = "<B><FONT FACE=corflaticon COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"


  Strsql = "Select * From OBJETOS  where CODOBJETO = " & sCODPROCESO 
  Strsql = Strsql & " and tipobjeto = '" & sTIPPROCESO & "'  "

 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     WKPROCDESC = RS(0)

      sHref = "javascript:void(0);"

      sImage = "<font face='corflaticon' size=2 color='#FFFFFF'>X</font>"
      sLink = ofv.GenLink("","nt onclick='document.all.DX4.style.display:""none"";return false;' ",sHref,HrfTrg,"","",sImage,"si")


       Filan = ""
       Filan1 = ""
       Celda = ""
       Celda = Celda & ofv.GenCelda("","ut COLSPAN=2 style='background-color:#993333;color:#ffffff;' ","","","","","","PROPIEDADES","si")  
       Celda = Celda & ofv.GenCelda("","aaa style='background-color:#993333;color:#ffffff;' onclick=""document.all.DX4.style.display='none';return false;"" ","center","","","","",sImage,"si")  
       FilaN = ofv.GenRow("","aaa bgcolor='#fefefe'","","","15","","",Celda,"si")

       wkproccod = "<a class=nt href='emdetaproc.asp?Nombre=" & request.querystring("Nombre")
       wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & rs(0) & "&tipmaster=" & rs(2) & "' ><b>"
       wkproccod = wkproccod & rs(3)
       wkproccod = wkproccod & "</b></a>"

       Celda = ""
       sInput = "&nbsp;"
       Celda = Celda & ofv.GenCelda("","UT ","","10%","","","",sInput,"si")
       Celda = Celda & ofv.GenCelda("","UT ","","89%","","","",sInput,"si")
       Celda = Celda & ofv.GenCelda("","UT ","","1%","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","15","","",Celda,"si")


       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","NOMBRE&nbsp;","si")
       sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"50","dt","")
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","CODIGO&nbsp;","si")
       sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"50","dt","")
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")


       Celda = ""
       scomK = "SP" & chr(9) & "VN"
       scomD = "PROCESO" & chr(9) & "EVENTO"

       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","TIPO&nbsp;","si")
       sInput = ofv.gencomboFX(sForm,rs(17).name,rs(17),scomK,scomD," ")
       Celda = Celda & ofv.GenCelda("","vt colspan=2   ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

''''''''''''''''

       IF (RS(17) & "") = "VN" THEN

       sInstrumento = "Sin documento asociado ........"

       if  rs(15) > 0 then

           Strsql = "SELECT * From objlnk where objid = " & rs(15)
           set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rs1.eof then
              wtipoL   = clng(rs1(6))
              wtipOBJ  = rs1(7)
              wUrl     = rs1(8)
              wUrlDesc = rs1(9)
              wtipo    = rs1(10)


              if trim(wtipo) = "LRGVN" then

                 Strsql = "Select xdescri From csGenObjetos where codobjeto = " & wtipoL
     ''      response.write strsql & "<br>"
                 set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
                 call ofv.cerrarconsulta(rsInstrumento)

              end if


          end if
          call ofv.cerrarconsulta(rs1)




       end if
       sInput = sInstrumento


         sID = "EdGENRsk.asp?"
         sID = sID & "ID=" & rs(0) & "&TABLA=OBJRSKREL&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta2 & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo



        sinput2 = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinput2 = sinput2  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinput2 = sinput2  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinput2 = sinput2  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinput2 = sinput2  & "none" & chr(34) & ";' "
        sinput2 = sinput2  & "  href='" & sID & "'  "
        sinput2 = sinput2  & "  TITLE='Asociar/cambiar documento operativo.'  ><b>&nbsp;" & sImagenI3 & "...&nbsp;</b></a>" 

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","EVENTO&nbsp;","si")

       Celda = Celda & ofv.GenCelda("","vtg colspan=2 ","","","","","",sInput & sinput2,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

      END IF

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","TEST&nbsp;","si")
       sInput = "<input type='button' value='test' class='Bt' ONCLICK='alert(document.EDPROCF.action);' >"
       Celda = Celda & ofv.GenCelda("","vtg colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")



'''''''''''''''''

    '   sAccion = "asp/actualizar.asp?fname=" & rs(0)
    '   sAccion = sAccion & "&tabla=OBJETOS&volver=" & vuelta
    '   Formulario = ofv.genform(sForm,"",sAccion,"fesquemas","post","",Filan,"si")


       Filan = Filan1 & ofv.GenTabla("","aaa ","","95%","","0","","0","1","","",FilaN,"si")

       sHTML = sHTML & "<CENTER>" & Filan & "</CENTER>"
       sHTML = sHTML & "</body></html>"

  ELSE


  end if


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

end if

response.write sHTML

%>

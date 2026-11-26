<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
      Generar = ""
      'response.buffer=false
      Set cn = OFV.conectar(OFV.strconn2)
      Set cn2 = OFV.conectar(OFV.strconn0)

      nID = ""
      sSector = ""
      nSeg = ""

      nID = Request.querystring("ID")

      sUN = request.querystring("UN")
      'response.write sUN
      IF sUN <> "" THEN
         SESSION("SUN") = SUN
      ELSE
         SUN = SESSION("SUN")
      END IF

XXXX = NEWEMPRE("NO")

    session("ASP")="SUBCONSULTA.ASP"

    sUN = request.querystring("EMPRE")
  'response.write "sEMPRE: " & sEMPRE
     IF sUN = ""  AND SESSION("UN") = "" THEN

       sUN = ATREMP(0, 0)
       SESSION("UN") = sUN & ""
       If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
     elseIF sUN <> "" THEN
      XXXX = NEWEMPRE(sUN)
      SESSION("UN")=sUN
     ELSE
      XXXX = NEWEMPRE(SESSION("UN"))
      sUN=SESSION("UN")
   end if




          fila  = ""
          filam = ""
  sHTML = ofv.FormHeader(session("FormN"))
             Dim celdam, filam, tablam

  TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


'shtml = shtml & tablah & "<br>"
''''''''''''''''''''''''



      xtseg = clng("0" & request.form("segm"))
      xtagr = clng("0" & request.form("agrm"))

      if xtseg = 0 then xtseg = 1



      Celda = ofv.GenCelda("","ttt colspan=2 ","","","","","","SELECCIONE EL SECTOR","si")
   '   Celda = Celda & ofv.GenCelda("","vti","","60%","","","","&nbsp;","si")
      Fila  = Fila & ofv.GenRow("","","","","50","","",Celda,"si")
'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","Site/Sucursal","si")

      Strsql = "Select * From sites where siteun = " & sun & "  Order By descri"
      sInput = ofv.GenerarCombo("Organigrama1","segm",xtseg,Strsql,cn,0,1,"")

      sAccion = "OrganigramaUn.asp"
      formx   =  ofv.GenForm("Organigrama1","",sAccion,"","post","",sInput,"si")
      Celda   =  Celda & ofv.GenCelda("","vti","","","","","",formx,"si")

      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



'-----------------------------
      Celda   = ofv.GenCelda("","tt","","","","","","Nivel","si")


      Strsql  = "Select distinct niveljerar.codnivel, niveljerar.descri  From eco,niveljerar,ECOSITE "
      STRSQL  = STRSQL & " where eco.ecoun = " & sun & " and eco.econiv = niveljerar.codnivel and "
      STRSQL  = STRSQL & " eco.eco = ECOSITE.ECO and ecoSITE.ECOSITE = " & xtseg & " Order By niveljerar.descri "
   '  response.write strsql & "<br>"
       scomKK = ""
    scomDD = ""
  ''  response.write strsql & "<br>" & chr(13)
    set rsniv = ofv.crearconsultaEx(Strsql,cn,1,parametros)
    DO UNTIL rsniv.eof
       scomKK = scomKK & rsniv(0) & chr(9)
       scomDD = scomDD & rsniv(1) & chr(9)
       RSNIV.MOVENEXT
    LOOP
    call ofv.cerrarconsulta(rsniv)

  sInput = ofv.gencomboFX("Organigrama2","agrm",xtagr,scomKK,scomDD,"")
       ' sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2).value,scomKK,scomDD,"")
      'sInput = ofv.GenerarCombo(sForm,rs(2).name,rs(2).value,StrSql1,cn3,0,1,"")
      'sInput  = ofv.GenerarCombo("Organigrama2","agrm",xtagr,Strsql,cn,0,1,"")
      sInput  = sinput & ofv.GenerarInput("segm",xtseg,"hidden","Organigrama2","1","dt","")

      sAccion = "OrganigramaUn.asp"
      formx   =  ofv.GenForm("Organigrama2","",sAccion,"","post","",sInput,"si")
      Celda   =  Celda & ofv.GenCelda("","vti","","","","","",formx,"si")
      Fila    =  Fila & ofv.GenRow("","","","","","","",Celda,"si")



'-----------------------------

      Celda    = ofv.GenCelda("","tt","","","","","","Area-Sector","si")


      Strsql   = "Select * From eco, ECOSITE where ecoun = " & sun & " and econiv = " & xtagr
      STRSQL   = STRSQL & " and eco.eco = ECOSITE.ECO and ECOSITE.ecosite = " & xtseg & " Order By descri"
   '   response.write strsql
    '  sInput   = ofv.GenerarCombox3(strSQL, cn, 0,1, "ListaP")
     ' sInput  = ofv.GenerarCombo("filler","ListaP",1,Strsql,cn,0,1,"no")

           '  scomK = "0" & chr(9)
     	   '  scomD = "SIN DEFINIR" & chr(9)

     	   '  strsql="select * from eco order by eco.DESCRI"

     	     set rseco = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     	     do while not rseco.eof
     	     scomK=scomK & rseco(0) & chr(9)
     	     scomd=scomd & rseco(1) & chr(9)
     	     rseco.movenext
     	     loop

     	      ssino = ""

     	     if ssino = "" then ssino = "0"

	   sInput = ofv.gencomboFX("filler","ListaP",ssino,scomK,scomD,"no")


      if xtseg = 0 then xtseg = "nsnc"
      if xtagr = 0 then xtagr = "nsnc"

      sInput   = sinput & ofv.GenerarInput("SECSITE",xtseg,"hidden","Organigrama","1","dt","")
      sInput   = sinput & ofv.GenerarInput("SECNIVEL",xtagr,"hidden","Organigrama","1","dt","")
      sAccion  = "OrganigramaUnGraf.asp"
      formx    = ofv.GenForm("Organigrama","",sAccion,"_blank","post","",sInput,"si")
     ' formx    = ofv.GenForm("Organigrama","",sAccion,"","post","",sInput,"si")

      Celda   =  Celda & ofv.GenCelda("","vti","","","","","",formx,"si")
      Fila     = Fila & ofv.GenRow("","","","","","","",Celda,"si")

      Celda = ofv.GenCelda("","ut","","30%","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","vti","","60%","","","","&nbsp;","si")
      Fila  = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


      Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")

      sInput   = ofv.GenerarInput("","Graficar","button","","10","bt onclick='validate();' ","")
      formx    = ofv.GenForm("FILLER","","","","post","",sInput,"si")
      Celda    =  Celda & ofv.GenCelda("","vti","","","","","",formx,"si")
      Fila     = Fila & ofv.GenRow("","","","","","","",Celda,"si")

     sHTML = sHTML & "<center>"
     sHTML = sHTML & OFV.gentabla("", "aaa", "", "80% ", "", "0", "0", "4", "0", "", "", fila, "si")

   ''  sHTML = sHTML & "<center>" & tablam  & "</center>"


    sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegMover,CantRegMostrar,"no","no",posi,"","","no",Agregar,"no","")

      Call OFV.cerrarconn(cn)

     sHTML = sHTML  & "</body></html>"


end if



Response.Write sHTML

function selecEmpre()

sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

   IF  Session("EMPMUN") = "S" THEN

   Fila    = ""

   SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

   Filaz    = ""
   Celdaz   = ""
   sHref   = ""

   sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
   sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
   Celdaz   = Celdaz & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
   Celdaz   = Celdaz & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
   Celdaz   = Celdaz & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

   Filaz    = filaz & ofv.GenRow("","","top","","","","",Celdaz,"si")

   Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Filaz,"si")





   shtml=shtml & "<script language=javascript>" & chr(13)
   shtml=shtml & "function shwDem(){" & chr(13)
   shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
   shtml=shtml & "}" & chr(13)
   shtml=shtml & "</script>" & chr(13)
 '----------------------------------------------
   sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
   sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
   TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


   TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
   TABLAX = TablaX & "<tr><td  "
   TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
   TABLAX = TablaX & "<tr><td colspan=3 "
   TABLAX = TablaX & " >&nbsp;</td></tr>"

   FOR IAEMP=1 TO TAEMP

    '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

      IF ATREMP(IAEMP,2) = 1 THEN

       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
       ELSE
          TABLAX = TablaX & "<tr>"
       END IF

       TABLAX = TablaX & "<td width=5 "
       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
       ELSE
          TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
       END IF

       sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
       TABLAX = TablaX & "<td width=5 "
       TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

       TABLAX = TablaX & "<td class=vt  >"

       'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
 		sHref = "OrganigramaUN.asp?EMPRE=" & ATREMP(IAEMP,0)
     '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
       sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
       TABLAX = TablaX & sLink
       TABLAX = TablaX & "</td>"
       TABLAX = TablaX & "</tr>"
      END IF

   NEXT

   TABLAX = TablaX & OFV.genFtabla()

   TABLA = Tabla & TABLAX

   TABLA = Tabla & "</DIV>"

   sHTML = sHTML & TABLA

  end if

end function

%>

<SCRIPT LANGUAGE="JavaScript">

function validate(){
  var y=document.getElementById("ListaP0").options

if (document.getElementById("ListaP0").selectedIndex==0 ||document.getElementById("ListaP0").selectedIndex==-1 )
{
	alert("Por favor ingrese Area-Sector");

	return false;
	}

	document.Organigrama.submit();
	return true;

}
//-->
</SCRIPT>

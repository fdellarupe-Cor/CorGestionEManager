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
      IF sUN <> "" THEN
         SESSION("SUN") = SUN
      ELSE
         SUN = SESSION("SUN")
      END IF






          fila  = ""
          filam = ""
  sHTML = ofv.FormHeader(session("FormN"))
             Dim celdam, filam, tablam


''''''''''''''''''''''''



      xtseg = clng("0" & request.form("segm"))
      xtagr = clng("0" & request.form("agrm"))

      if xtseg = 0 then xtseg = 1



      Celda = ofv.GenCelda("","ttt colspan=2 ","","","","","","seleccione el sector","si")
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
   ''   response.write strsql & "<br>"
      sInput  = ofv.GenerarCombo("Organigrama2","agrm",xtagr,Strsql,cn,0,1,"")
      sInput  = sinput & ofv.GenerarInput("segm",xtseg,"hidden","Organigrama2","1","dt","")

      sAccion = "OrganigramaUn.asp"
      formx   =  ofv.GenForm("Organigrama2","",sAccion,"","post","",sInput,"si")
      Celda   =  Celda & ofv.GenCelda("","vti","","","","","",formx,"si")
      Fila    =  Fila & ofv.GenRow("","","","","","","",Celda,"si")



'-----------------------------

      Celda    = ofv.GenCelda("","tt","","","","","","Area-Sector","si")


      Strsql   = "Select * From eco, ECOSITE where ecoun = " & sun & " and econiv = " & xtagr
      STRSQL   = STRSQL & " and eco.eco = ECOSITE.ECO and ECOSITE.ecosite = " & xtseg & " Order By descri"
      sInput   = ofv.GenerarCombox3(strSQL, cn, 0,1, "ListaP")

      if xtseg = 0 then xtseg = "nsnc"
      if xtagr = 0 then xtagr = "nsnc"



''''''''''''''''

      XENVIADOR = ""

      STRSQL  = ""
      STRSQL  = STRSQL & " select b.eco, b.econiv, b.ecojerar, b.ecoun, b.ecosite,a.*, c.*, b.* from ecorel a"
      STRSQL  = STRSQL & " inner join eco b on (a.ecomaster = b.eco)"
      STRSQL  = STRSQL & " inner join niveljerar c on (b.econiv = c.codnivel and c.codnivel > 0)"
      STRSQL  = STRSQL & " order by codnivel, ecojerar"

      SET RS=ofv.crearconsultaEx(StrSql,cn,1,volver)
      IF not rs.eof then
         xtseg   = RS(4)
         xtagr   = RS(1)
         XListaP = RS(0)

         sInput   = ofv.GenerarInput("ListaP",XListaP,"hidden","Organigrama","1","dt","")

         XENVIADOR = " <SCRIPT LANGUAGE='JavaScript' >  document.Organigrama.submit();  </SCRIPT>"


      END IF


'''''''''''''''
      sInput   = sinput & ofv.GenerarInput("SECSITE",xtseg,"hidden","Organigrama","1","dt","")
      sInput   = sinput & ofv.GenerarInput("SECNIVEL",xtagr,"hidden","Organigrama","1","dt","")
      sAccion  = "OrganigramaUnGraf.asp"

     ' formx    = ofv.GenForm("Organigrama","",sAccion,"_blank","post","",sInput,"si")
      formx    = ofv.GenForm("Organigrama","",sAccion,"","post","",sInput,"si")

      Celda   =  Celda & ofv.GenCelda("","vti","","","","","",formx,"si")
      Fila     = Fila & ofv.GenRow("","","","","","","",Celda,"si")

      Celda = ofv.GenCelda("","ut","","30%","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","vti","","60%","","","","&nbsp;","si")
      Fila  = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


      Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")

      sInput   = ofv.GenerarInput("","GRAFICAR","button","","10","bt onclick='validate();' ","")
      formx    = ofv.GenForm("FILLER","","","","post","",sInput,"si")
      Celda    =  Celda & ofv.GenCelda("","vti","","","","","",formx,"si")
      Fila     = Fila & ofv.GenRow("","","","","","","",Celda,"si")

     sHTML = sHTML & "<center><br><br>"
     sHTML = sHTML & OFV.gentabla("", "aaa style='display:none;border-style:solid;border-width:1;border-color:#eeeeee;' ", "", "40% ", "", "0", "0", "4", "0", "", "", fila, "si")

   ''  sHTML = sHTML & "<center>" & tablam  & "</center>"


    sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegMover,CantRegMostrar,"no","no",posi,"","","no",Agregar,"no","")

      Call OFV.cerrarconn(cn)


     sHTML = sHTML  & XENVIADOR

     sHTML = sHTML  & "</body></html>"


end if

Response.Write sHTML



%>

<SCRIPT LANGUAGE="JavaScript">

function validate(){
if (document.getElementById("ListaP0").selectedIndex==0 ||document.getElementById("ListaP0").selectedIndex==-1 )
{
	alert("Por favor ingrese Area-Sector");
	
	return false;
	}

	EnviaForm();
	return true;

}
</SCRIPT>

<SCRIPT LANGUAGE="JavaScript">
function EnviaForm(){

document.Organigrama.submit();

}
//-->
</SCRIPT>

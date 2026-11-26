
<%

 FUNCTION GENFMTFAQ(FNCODOBJ,FNTIPOBJ)


  FNshtml = ""

  Strsql = "SELECT * FROM OBJFAQ  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
''response.write strsql & " strsql " & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then

     xsurl = rs(3)

 
   if instr(ucase(xsurl),"HTTP://") = 0 then xsurl = "HTTP://" & xsurl 
  
      FNshtml = FNshtml & "<form id=dpsch action='" & xsurl & "' method='post' target='_blank' >"

    If IsNull(rs(5)) and IsNull(rs(6)) and IsNull(rs(7)) and IsNull(rs(8)) and IsNull(rs(9)) and IsNull(rs(10)) Then 

     Whtml = "<CENTER><br><br>"
     Whtml = Whtml & "<DIV id='PDHLP' class='DINDM'  STYLE='display:block;z-index:99;width:350;margin-top:1;margin-left:1;"
     Whtml = Whtml & "'>"
     Whtml = Whtml & "<table id=""MTHLP1""  class='TWIND' width=350 cellspacing=0 cellpadding=2 >"
     Whtml = Whtml & "<tr>"
     Whtml = Whtml & "<td class=mbt width=99% >&nbsp;ATENCION</td>"
     Whtml = Whtml & "</tr>"
     Whtml = Whtml & "</table>"
     Whtml = Whtml & "<table id=""MTHLP2""  class=aaa width=350  cellspacing=0 cellpadding=6 "
     Whtml = Whtml & " style=""border-style:solid;border-width:3;border-color:#eeeeee;background-color:#f9f9f9;"" "
     Whtml = Whtml & " >"
     Whtml = Whtml & "<tr HEIGHT=100 >"
     Whtml = Whtml & "<td class=UT width=100% style='text-align:justify;' >EN CASO DE NO VISUALIZAR LA VENTANA EMERGENTE "
     Whtml = Whtml & "POR FAVOR PRESIONE EL BOTON ""VISUALIZAR"" </td>"
     Whtml = Whtml & "</tr>"
     Whtml = Whtml & "<tr HEIGHT=40 >"
     Whtml = Whtml & "<td class=UT align=center valign=top >"
     Whtml = Whtml & "<input type=submit value='VISUALIZAR' CLASS=BT >"
     Whtml = Whtml & "</td>"
     Whtml = Whtml & "</tr>"
     Whtml = Whtml & "</table>"
     Whtml = Whtml & "</DIV>"

     FNshtml = FNshtml & Whtml

    end if

     FNshtml = FNshtml & "<input type=hidden value=dspch  ></form></CENTER>"

   
   
     FNshtml = FNshtml & "<script language=javascript>"
     FNshtml = FNshtml & "dpsch.submit();"
     FNshtml = FNshtml & "</script>"

 
     
  ELSE


     FNshtml = FNshtml & "<BR><BR><BR><P STYLE='COLOR:#993333;FONT-FAMILY:VERDANA;FONT-SIZE:9pt;'><B>NO TIENE DEFINIDA UNA DIRECCION PARA PREGUNTAS FRECUENTES.</P> "


  end if

  GENFMTFAQ = FNshtml

END FUNCTION

%>
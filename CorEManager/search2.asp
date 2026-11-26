<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="configuration.asp" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  set cn = ofv.conectar(ofv.strconn2)


  sHTML = ofv.FormHeader(session("FormN"))

 %>
 <script language="JScript" runat="server">

 function Ceil(number)
 {
   return Math.ceil(number)
 }

 function Floor(number)
 {
   return Math.floor(number)
 }

 function TheMaximum(number1, number2)
 {
   return Math.max(number1, number2)
 }

 function TheMinimum(number1, number2)
 {
   return Math.min(number1, number2)
 }

 </script>

<script language="JScript">
function AbrirCons(Datos,nom) {

 var h=570;w=790 ;
 if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;}
 myWin= open(Datos,'ConsLnk' + nom,  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes' );
}

</script>
 <%


    'on error resume next

    set Csearch = server.createobject("corsearch.search")   'new CYiderSearch
    set Csearch.ofv=ofv
    set Csearch.response=response
    set Csearch.server=server

    g_set = g_set + 1

    Csearchh.m_charset = g_charset
    Csearch.Constructor g_database_connection

    Csearch.m_style_title = g_style_title
    Csearch.m_style_text = g_style_text
    Csearch.m_style_you_searched_for = g_style_you_searched_for
    Csearch.m_style_url = g_style_url
    Csearch.m_style_br = g_style_br
    Csearch.m_style_more_results_text = g_style_more_results_text
    Csearch.m_style_more_results_link = g_style_more_results_link
    Csearch.m_style_more_results_link_Next = g_style_more_results_link_Next
    Csearch.m_trailing_words = g_trailing_words

    Csearch.DisplayResults
    Csearch.Destructor
    set yCsearch = Nothing
  g_set = g_set - 1

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
shtml = "<html>"

shtml = shtml & "<head>"
shtml = shtml & "<meta charset='utf-8'>"
shtml = shtml & "<meta name='viewport' content='width=device-width, initial-scale=1'>"
shtml = shtml & "<title>jQuery test Mauro</title>"
shtml = shtml & "<link rel='stylesheet' href='//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css'>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery.min.js'></script>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery-ui.js'></script>"
shtml = shtml & "</head>"

shtml = shtml & "<body>"

shtml = shtml & "<div id='accordion' style='width:400px;'>"
shtml = shtml & "  <h3>Section 1</h3>"
shtml = shtml & "  <div>"
shtml = shtml & "    <p>Mauris mauris ante, blandit et, ultrices a, suscipit eget."
shtml = shtml & "    Integer ut neque. Vivamus nisi metus, molestie vel, gravida in,"
shtml = shtml & "    condimentum sit amet, nunc. Nam a nibh. Donec suscipit eros."
shtml = shtml & "    Nam mi. Proin viverra leo ut odio.</p>"
shtml = shtml & "  </div>"
shtml = shtml & "  <h3>Section 2</h3>"
shtml = shtml & "  <div>"
shtml = shtml & "    <p>Sed non urna. Phasellus eu ligula. Vestibulum sit amet purus."
shtml = shtml & "    Vivamus hendrerit, dolor aliquet laoreet, mauris turpis velit,"
shtml = shtml & "    faucibus interdum tellus libero ac justo.</p>"
shtml = shtml & "  </div>"
shtml = shtml & "  <h3>Section 3</h3>"
shtml = shtml & "  <div>"
shtml = shtml & "    <p>Nam enim risus, molestie et, porta ac, aliquam ac, risus."
shtml = shtml & "    Quisque lobortis.Phasellus pellentesque purus in massa.</p>"
shtml = shtml & "    <ul>"
shtml = shtml & "      <li>List item one</li>"
shtml = shtml & "      <li>List item two</li>"
shtml = shtml & "      <li>List item three</li>"
shtml = shtml & "    </ul>"
shtml = shtml & "  </div>"
shtml = shtml & "</div>"

shtml = shtml & "<script>"
shtml = shtml & "$( '#accordion' ).accordion({"
shtml = shtml & "  heightStyle: 'fill'"
shtml = shtml & "});"
shtml = shtml & "</script>"

shtml = shtml & "</body>"
shtml = shtml & "</html>"

response.write shtml
%>

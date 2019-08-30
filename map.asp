<%@ Language="VBScript" %>
<% Response.Buffer=true  
  Response.CharSet = "big5"
  Session.codepage = 950   

  nowprogram_ls = "map" 
  Program_name = "地圖"  
%> 
<%
//2013-03-26  增加多語系

Pigeon_site_no = trim(Request("Pigeon_site_no"))
ringno = trim(Request("ringno"))
msgdatetime=trim(Request("backtime"))
lo=trim(Request("lo"))
lo_ls_w = mid(lo,1,1)
la=trim(Request("la"))
la_ls_s = mid(la,1,1)
GPSlang = trim(Request("GPSlang"))

if lo<>"" then ee_no=int(mid(lo,2,3)) + (int(mid(lo,5,2)) *60+int(mid(lo,7,2)))/3600
if la<>"" then nn_no=int(mid(la,2,3)) + (int(mid(la,5,2)) *60+int(mid(la,7,2)))/3600

show_lo = mid(lo,2,3) & "&#176;" & mid(lo,5,2) & "'" & mid(lo,7,2) 
show_la = mid(la,2,3) & "&#176;" & mid(la,5,2) & "'" & mid(la,7,2)

e=Round(ee_no,6)
n=Round(nn_no,6)



  if lo<>"" and lo_ls_w = "W" then e1= "W" Else e1 = "E" End if
  if la<>"" and la_ls_s = "S" then n1= "S" Else n1 = "N" End if
  showmsg_ls = "Loft: " & Pigeon_site_no &"<br>Ring: " & ringno & "<br>Arrival Time: " & msgdatetime & "<br>Longitude : " & e1 & show_lo & "<br>Latitude : " & n1 & show_la

%>
<!DOCTYPE html>
<html lang="en">
<head>

    <meta charset="big5">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="TOPigeon Pigeon Clock System">
    <meta name="author" content="">
	
	<link rel="shortcut icon" href="images/icon_32.ico">	
	<title>TOPigeon</title>

</head>
    <script src="https://maps.google.com/maps?file=api&v=2&key=ABQIAAAAG_4i2swR3KOd-nGYrlrt8RTkyS8SRe_kYPTAbwTumvAqao01PRRUcCtCzTBnNH2kRURGR8RhQQoZ3w" type="text/javascript"></script>

    <script type="text/javascript">

    function initialize() {
      if (GBrowserIsCompatible()) {
        var map = new GMap2(document.getElementById("map_canvas"));
        var center = new GLatLng(<%=n%>,<%=e%>);
        map.setCenter(center, 15);
        map.setUIToDefault();
        
        var myIcon = new GIcon();
        myIcon.image = "images/a30101.gif"; 
	myIcon.iconAnchor = new GPoint(13, 20); 
	myIcon.infoWindowAnchor = new GPoint(13, 0); 
	myIcon.iconSize = new GSize(35, 35); 
        
        var marker = new GMarker(center, { icon:myIcon, draggable: false});
        map.addOverlay(marker);
        marker.openInfoWindowHtml( "<%=showmsg_ls%>" );

        GEvent.addListener(marker, "dragstart", function() {
          map.closeInfoWindow();
        });

        GEvent.addListener(marker, "dragend", function() {
          marker.openInfoWindowHtml("Just bouncing along...");
        });
        
      }
    }
    </script>
    
  <body onload="initialize()" onunload="GUnload()">
    <div id="map_canvas" style="width: 500px; height: 300px"></div>
  </body>
</html>



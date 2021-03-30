<%@ page language="java" contentType="text/html; charset=ISO-8859-1" pageEncoding="ISO-8859-1"%>
<%@page import="java.io.OutputStream"%>   
<%@ page import="java.io.*"%>
<%@ page import="java.util.*"%>
<%@ page import="javax.servlet.*"%>

<%@ page import="com.olympus.olyutil.Olyutil"%> 

<%   	String title =  "Olympus FIS Asset Upload Report Error Page"; 
String dateErr =  (String) session.getAttribute("dateErr");
	
String idErr =  (String) session.getAttribute("idErr");
 
	
%>

<!DOCTYPE html>
<html>
<head>
<meta charset="ISO-8859-1">
<title><%=title%></title>
<script type="text/javascript" src="http://cdnjs.cloudflare.com/ajax/libs/jquery/1.9.1/jquery.min.js"></script>
<script type="text/javascript" src="http://cdnjs.cloudflare.com/ajax/libs/jquery.tablesorter/2.9.1/jquery.tablesorter.min.js"></script>
<script type="text/javascript" src="includes/js/tableFilter.js"></script>

<style><%@include file="includes/css/header.css"%></style>
 
    <link rel="stylesheet" href="includes/css/calendar.css" /> 
<style><%@include file="includes/css/table.css"%></style>




<!-- ******************************************************************************************************************************************************** -->
<style>

</style>
<!-- ******************************************************************************************************************************************************** -->
<script>

$(function() {

  // call the tablesorter plugin
  $("table").tablesorter({
    theme: 'blue',
    // initialize zebra striping of the table
    widgets: ["zebra"],
    // change the default striping class names
    // updated in v2.1 to use widgetOptions.zebra = ["even", "odd"]
    // widgetZebra: { css: [ "normal-row", "alt-row" ] } still works
    widgetOptions : {
      zebra : [ "normal-row", "alt-row" ]
    }
  });

});	
			
    </script>
    
   <script>
  function openWin(myID) {
  
  
   myID2 = document.getElementById(b_app).value;

  alert("ID" + myID2);
  //window.open("http://cvyhj3a27:8181/fisAssetServlet/readxml?appID=" + myID2);
	}
	
	
	var call = function(id){
		var myID = document.getElementById(id).value;
		//alert("****** myID=" + myID + " ID=" + id);
		//window.open("http://cvyhj3a27:8181/fisAssetServlet/readxml?appID=
		window.open("http://localhost:8181/webreport/getquote?appKey=" + myID);
				
				
	}



	var getExcel = function(urlValue){
		var formUrl = document.getElementById(urlValue).value;
		//alert("SD=" + startDate + "****** formUrl=" + formUrl + " \n***** urlValue=" + urlValue);
		//alert("in Quote" + myID + " --- id=" +id);
		window.open( formUrl, 'popUpWindow','height=500,width=800,left=100,top=100,resizable=yes,scrollbars=yes,toolbar=yes,menubar=no,location=no,directories=no, status=yes' );
	
	}
	
	
	
</script> 
</head>
<!-- ********************************************************************************************************************************************************* -->

<body>

 
<%@include  file="includes/header.html" %>

<h3><%=title%></h3>

<%!

 
	
	%>



 
<%
 	/*****************************************************************************************************************************************************************/

 %>
	
   
   
   
			<div style="height: 450px; overflow: auto;">
		<% 			
		//out.println("<br><h5>idErr="   +  idErr +    "-- </h5><br>");
		//out.println("<br><h5>dateErr="   +  dateErr +    "-- </h5><br>");

	 
		
		if (! Olyutil.isNullStr(dateErr) && dateErr.equals("dateErr")) {
			out.println("<br><h5>An Error occurred. Please re-submit vith a valid Date. (ex. 2021-03-19) </h5><br>");

		} else if (! Olyutil.isNullStr(idErr) & idErr.equals("idErr")) {
			out.println("<br><h5>An Error occurred. Please re-submit vith a valid Contract ID. (ex. 101-0008740-006) </h5><br>");
		}
		/**********************************************************************************************************************************************************/
		// Output Table 
	
	  request.getSession().setAttribute("dateErr", "");
	  request.getSession().setAttribute("idErr", "");
		
%>	
</div>
		

 



</body>
</html>
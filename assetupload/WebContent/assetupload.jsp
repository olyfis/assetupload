<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
    <%@page import="java.io.OutputStream"%>   
<%@ page import="java.io.*"%>
<%@ page import="java.util.*"%>
<%@ page import="javax.servlet.*"%>
    
    
    <% 
	String title =  "Olympus Asset Upload File Search"; 	 
	ArrayList<String> tokens = new ArrayList<String>();
	String formUrl =  (String) session.getAttribute("formUrl");
%>
    
    
    
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
 
<!--  	
http://cvyhj3a27:8181/nbvaupdate/nbvaupdate?id=101-0010311-004&eDate=2020-03-10&boDate=2020-03-31&invoice=123123123

101-0009442-019-->

<title><%=title%></title>


<!--  <link href="includes/appstyle.css" rel="stylesheet" type="text/css" /> 

<style><%@include file="includes/css/reports.css"%></style>
<style><%@include file="includes/css/table.css"%></style>
-->
<style><%@include file="includes/css/header.css"%></style>
<style><%@include file="includes/css/menu.css"%></style>
 <link rel="stylesheet" href="includes/css/calendar.css" />
    <script type="text/javascript" src="includes/scripts/pureJSCalendar.js"></script>


 

<script language="javascript" type="text/javascript">
 
//Browser Support Code
function ajaxFunction(){
	var ajaxRequest;  // The variable that makes Ajax possible!
	
	try{
		// Opera 8.0+, Firefox, Safari
		ajaxRequest = new XMLHttpRequest();
	} catch (e){
		// Internet Explorer Browsers
		try{
			ajaxRequest = new ActiveXObject("Msxml2.XMLHTTP");
		} catch (e) {
			try{
				ajaxRequest = new ActiveXObject("Microsoft.XMLHTTP");
			} catch (e){
				// Something went wrong
				alert("Your browser broke!");
				return false;
			}
		}
	}
	// Create a function that will receive data sent from the server
	ajaxRequest.onreadystatechange = function(){
		if(ajaxRequest.readyState == 4){
			var ajaxDisplay = document.getElementById('ajaxDiv');
			ajaxDisplay.innerHTML = ajaxRequest.responseText;
		// document.actionform.actiontypeprogram.value = ajaxRequest.responseText;

		}
	}

//	var atype = document.actionform.getElementById('actiontype').value;
	var atype = document.actionform.actiontype.value;
	var date2 = document.actionform.startDate.value;
	//alert("atype=" + atype);
	
	//alert("date2=" + date2);
	//var queryString = "?atype=" + atype + "&wpm=" + wpm + "&ex=" + ex;
	var queryString = "/webreport/ajaxIL.jsp?atype=" + atype + "&date2=" + date2;
	

	  
	//alert("QS=" + queryString);
	//ajaxRequest.open("GET", "ajax.jsp" + queryString, true);
	ajaxRequest.open("POST", queryString, true);
	ajaxRequest.send(); 
}

 
</script>





<!-- ********************************************************************************************************************************************************* -->

</head>
<body>
    
    
 <%@include  file="includes/header.html" %>


        
<!--   <img src="includes\images\logo.jpg" alt="logo"  height="100" width="225" align="right"> -->


<div style="padding-left:20px">

</div>

<BR>
<h5>This page will provide an on-demand list of contract details from the Asset Upload spreadsheet.</h5>
<h5>Enter a valid contract ID.</h5>




 <h5>Note: <font color="red">This application will process a large data set and may take some time to run.</font> <BR>
 
<BR>
<!--   ************** Add Forms ********************************************************************************************************************* -->

<form name="actionform" method="get" enctype="multipart/form-data" action="assetup"  name="id"> 

<BR>


<table class="a" width="40%"  border="1" cellpadding="1" cellspacing="1">
<tr> <th class="theader"> Olympus FIS Asset Upload Query by Contract ID</th> </tr>
  <tr>
    <td class="table_cell">
    <!--  Inner Table -->
    <table class="a" width="100%"  border="1" cellpadding="1" cellspacing="1">
  <tr>
  <td width="20" valign="bottom"> <b>Enter Contract ID: (Ex. 101-0016020-001)</b> </td> 
  <td width="20" valign="bottom">  
     <%  out.println("<input name=\"id\" id=\"id\" type=\"text\" value=\"\"   />" );
     %>
  </td>
  </tr>
  

  <tr>
   <td  valign="bottom" class="a">
	<div id='ajaxDiv'> </div>
	</td>
	 <td> 
    <INPUT type="submit" value="Submit">  
    </td>
	
  </tr>
  </table>

</table>

 </form>
 <!-- ************ End First Table *************** -->
 <BR>
 <form name="actionform" method="get" action="assetup">

<BR>


<table class="a" width="40%"  border="1" cellpadding="1" cellspacing="1">
<tr> <th class="theader"> Olympus FIS Asset Upload Query by Invoice Date</th> </tr>
  <tr>
    <td class="table_cell">
    <!--  Inner Table -->
    <table class="a" width="100%"  border="1" cellpadding="1" cellspacing="1">
  <tr>
  <td width="20" valign="bottom"> <b>Enter Invoice Date: (Ex.2021-03-19)</b> </td> 
  <td width="20" valign="bottom">  
     <%  //out.println("<input name=\"appid\" id=\"appid\" type=\"text\" value=\"\"   />" );
      out.println("<input name=\"date2\" id=\"date2\" type=\"text\" value=\"Click for Calendar\" onclick=\"pureJSCalendar.open('yyyy-MM-dd', 20, 30, 7, '2017-1-1', '2025-12-31', 'date2', 20)\"   />" );
     
     %>
  </td>
  </tr>
  

  <tr>
   <td  valign="bottom" class="a">
	<div id='ajaxDiv'> </div>
	</td>
	 <td> 
    <INPUT type="submit" value="Submit">  
    </td>
	
  </tr>
  </table>

</table>
<!-- ************ End Second Table *************** -->
 </form>
 
 
 </BR>


<!--   ************** End  Add Forms **************************************************************************************************************** --> 

<h5>If you require access to the reports, please contact: John.Freeh@olympus.com</h5>
<h5>Note: <font color="red">Requires Javascript to be enabled.</font> <BR>
<script type="text/javascript" src="includes/js/validateID.js"></script>
</body>
</html>
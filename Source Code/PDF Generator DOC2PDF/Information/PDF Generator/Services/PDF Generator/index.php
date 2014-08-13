<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<link rel="stylesheet" type="text/css" href="/Commerce/Includes/Stylesheet/Commerce_Website.css">
<title>PDF Generator : Online Document Convertor</title>
<script lang="javascript">
function ValidateSubmit()
{
if (document.uploadForm1.emailAddress.value.length < 1)
{
alert('Please enter an Email Address. This is obviously important because else you will never learn about your converted document\'s whereabouts');
document.uploadForm1.emailAddress.focus();
return false;
}
else
{
var re = new RegExp('[\w\.-]*[a-zA-Z0-9_]@[\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]');
  if (document.uploadForm1.emailAddress.value.match(re)) {
    return true;
  } else {
alert('Please enter a valid Email Address. This is obviously important because else you will never learn about your converted document\'s whereabouts');
    return false;
  }


}
}
</script>
</head>

<body>
<table width="100%" height="100%" cellpadding="15" border="0"><tr><td valign="top">
<table width="100%"><tr><td align="left" width="50%">
	<a href="/Services/PDF%20Generator/">
	<img border="0" src="images/PDFGenerator.jpg" width="362" height="62"></a></td>
	<td align="right" width="50%" valign="top">
	<a href="http://www.commerce.uct.ac.za"><img border="0" src="images/Banner_Image.jpg" ></a>
</td></tr></table>
<table width="100%" cellspacing="0" cellpadding="4"><tr><td align="left" width="70%" valign="top" rowspan="2">

<?php


function handle_error($Error) {
	echo "An Error occured: ".$Error;
	//die();
}

function formatTime($sTime) {
	if (!is_numeric($sTime)) return false;
	
	$Year = (int) substr($sTime,0,4);
	$Month = (int) substr($sTime,4,2);
	$Day = (int) substr($sTime,6,2);
	$Hour = (int) substr($sTime,8,2);
	$Minute = (int) substr($sTime,10,2);
	
	if (!checkdate($Month, $Day, $Year)) return false;
	if ($Hour > 23) return false;
	if ($Minute > 59) return false;
	return mktime($Hour, $Minute, 0, $Month, $Day, $Year);
}



function parseString($sString) {
	$arrString = explode("/",$sString);
	return $arrString[count($arrString)-3]." :: ".$arrString[count($arrString)-2];
}

session_start();
if (isset($_SESSION['studentnumber']) == false)
{
//header('location:login.asp?str_username='.$_POST["username"].'&errorcase=99&site='. $_GET["site"] .'&URL='. $_GET["URL"]);							
}
//echo "<h1>Current User: " . $_SESSION['studentnumber'] . "</h1>\n";
?>
<h1><img border="0" src="images/OnlinePDFGenerator.jpg" width="400" height="45"></h1>
<p>Creating and distributing your work files as PDF documents has just become a 
whole lot easier. Simply fill in your email address and upload your Word or 
other format documents using the form below and click on the Upload button to 
continue. Online PDF Generator will then proceed to convert your document's 
contents into PDF format and once complete will mail you the URL from which you 
can download your new PDF file. </p>
<p>The generated PDF file will be available for download for 24 hours only, so 
be sure to grab it when you get your mail!</p>
<?php

$dir = "/Services/PDF Generator/Input";
$dir = "C:\\Inetpub\\wwwroot" . str_replace("/","\\",$dir);
if ($dir == "C:\\Inetpub\\wwwroot")
{
$dir = "";
}


if ($writtenstatus == false)
{
?>
<hr noshade color="#C0C0C0" size="1">
<h2>Upload your File for Conversion</h2>
<p>You can use the following upload form to post your file for conversion. <i><br>
<br>
(Please note that you cannot upload a folder)</i></p>
		<form enctype="multipart/form-data" action="postform.php" method="POST" name="uploadForm1" onsubmit="return ValidateSubmit();">
		<input type="hidden" name="URL" value="<?php echo $dir; ?>">
<div align="center">
<table border="0" cellpadding="4">
	<tr><td>File to upload:&nbsp; 
	<input name="userfile[]" type="file" size="55" class="boxBlur" onfocus="this.className='boxFocus';" onblur="this.className='boxBlur';"/></td></tr>
<!--<tr><td>File to upload:&nbsp; <input name="userfile[]" type="file" size="55" class="boxBlur" onfocus="this.className='boxFocus';" onblur="this.className='boxBlur';"/></td></tr>-->
<tr><td>Email Address: 
	<input name="emailAddress" type="text" size="55" class="boxBlur" onfocus="this.className='boxFocus';" onblur="this.className='boxBlur';"/>
	<font color="#FF0000">*required</font></td></tr>
<tr><td align="center">
	<input type="submit" value="Upload your file for conversion" /></td></tr></table>
		</div>  
		</form>

<hr noshade color="#C0C0C0" size="1">
		
<?php		
}							
	
?>		
</td>
	<td align="left" valign="top" bgcolor="#E2E5E7">
	<hr noshade color="#C0C0C0" size="1">
	<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="5">
		<tr>
			<td><p align="center"><b><font color="#FF8A16">:: NOTICE 
	::</font></b><font color="#FF8A16"></p><p>
	Please note that 
	any uploaded files that are suspected of carrying viruses will automatically 
	be removed by the server's anti-virus application. </font></p>
			</td>
		</tr>
	</table>
	<hr noshade color="#C0C0C0" size="1">
</td></tr>
	
<tr><td valign="bottom" align="center" bgcolor="#E2E5E7">
	&nbsp;</td>
	</tr>
	</table>
<!-- End Content -->
</td></tr>
<tr>
<td>
<p align="center">
	<table cellSpacing="0" cellPadding="0" border="0" id="table1">
		<tr align="middle">
			<td><a href="http://www.uct.ac.za">University of Cape Town</a> |
			<a href="http://www.commerce.uct.ac.za">Faculty of Commerce</a> |
			<a href="/disclaimer.asp">Disclaimer</a> |
			<a href="/Commerce/Information/contact_information.asp">
			Contact Faculty Office</a> |
			<a href="mailto:commercehelpdesk@uct.ac.za">IT Helpdesk</a> |
			<a href="/Commerce_IT">Managed by Commerce I.T.</a> </td>
		</tr>
		<tr align="middle">
			<td><font color="#999999" size="1"><strong>Copyright  <?php echo date("Y", date("U")); ?> Faculty 
			of Commerce -- University of Cape Town</strong></font></td>
		</tr>
	</table>
</p>
</td>
</tr></table>

</body>
</html>

<html>
<head>
<meta http-equiv="Content-Language" content="en-za">
<link rel="stylesheet" type="text/css" href="/Commerce/Includes/Stylesheet/Commerce_Website.css">
<title>PDF Generator : Online Document Convertor</title>
</head>

<body>
<?php
//mAXIMUM IDLE TIME = 15 MINUTES
set_time_limit(900);

error_reporting(0);

function parseString($sString) {
	$arrString = explode("/",$sString);
	return $arrString[count($arrString)-3]." :: ".$arrString[count($arrString)-2];
}

session_start();

if (isset($_SESSION['studentnumber']) == false)
{
//header('location:login.asp?str_username='.$_POST["username"].'&errorcase=99&site='. $_GET["site"] .'&URL='. $_GET["URL"]);							
}
?>

<table width="100%" height="100%" cellpadding="15" border="0"><tr><td valign="top">
<table width="100%"><tr><td align="left" width="50%">
	<a href="/Services/PDF%20Generator/">
	<img border="0" src="images/PDFGenerator.jpg" width="362" height="62"></a></td>
	<td align="right" width="50%" valign="top">
	<a href="http://www.commerce.uct.ac.za"><img border="0" src="images/Banner_Image.jpg" ></a>
</td></tr></table>
<table width="100%" cellspacing="0" cellpadding="4"><tr><td align="left" width="70%" valign="top" rowspan="2">



<h1><img border="0" src="images/OnlinePDFGenerator.jpg" width="400" height="45"></h1>
<h1 class="Error">The file Upload Process has completed.</h1>
<p>Should your upload have been successful, your document will be picked up and converted to PDF format shortly. An email will then 
be sent to the email address provided, informing you of the download location of this file.</p>
<p>The generated PDF file will be available for download for 48 hours only, so 
be sure to grab it when you get your mail!</p>
<p align="center"><a href="index.php">:: Click here to convert more files ::</a></p>


<?php

$dir = $_POST["URL"];
//$dir = "C:\\Inetpub\\wwwroot" . str_replace("/","\\",$dir);



//Path where the files will be uploaded to
$UploadPath = $dir . "\\";
//echo $dir . "\\";

//Track all the files uploaded successfully
$arrFiles = array();


try {
	if (!is_dir($UploadPath)) {
		handle_error("Upload location is not currently available");
		die();
	}


	//Check if the $_FILES variable exists
	if (!isset($_FILES['userfile'])) {
		handle_error("Upload cannot locate the file you want to upload");
		die();
	}

	for ($h = 0; $h < count($_FILES['userfile']['name']); $h++) {
		if ($_FILES['userfile']['name'][$h] == "") continue;
		
		//Check to see if an error occurred during the file upload

		$FileUploadError = $_FILES['userfile']['error'][$h];
		if ($FileUploadError != 0) {
			/*UPLOAD_ERR_OK
				Value: 0; There is no error, the file uploaded with success. 
				
				UPLOAD_ERR_INI_SIZE
				Value: 1; The uploaded file exceeds the upload_max_filesize directive in php.ini. 
				
				UPLOAD_ERR_FORM_SIZE
				Value: 2; The uploaded file exceeds the MAX_FILE_SIZE directive that was specified in the HTML form. 
				
				UPLOAD_ERR_PARTIAL
				Value: 3; The uploaded file was only partially uploaded. 
				
				UPLOAD_ERR_NO_FILE
				Value: 4; No file was uploaded. 
				
				UPLOAD_ERR_NO_TMP_DIR
				Value: 6; Missing a temporary folder. Introduced in PHP 4.3.10 and PHP 5.0.3. 
				
				UPLOAD_ERR_CANT_WRITE
				Value: 7; Failed to write file to disk. Introduced in PHP 5.1.0. 
			*/
				
			switch ($FileUploadError) {
				case 1:
					handle_error("The specified file is too large to be uploaded");
				break;			
				case 2:
					handle_error("The specified file is too large to be uploaded");
				break;
				case 3:
					handle_error("The specified file was not uploaded completely");
				break;
				case 4:
					handle_error("No file was specified to be uploaded");
				break;
				case 6:
					handle_error("Unable to locate the upload location");
				break;
				case 7:
					handle_error("Unable to write to the server");
				break;
			}
			die();
		}

		//Tells whether the file was uploaded via HTTP POST
		if (!is_uploaded_file($_FILES['userfile']['tmp_name'][$h])) {
			handle_error("Possible file upload attack");
			die();
		}
	
		//Remove special chars from the file
		$FileName = strip_File($_FILES['userfile']['name'][$h]);
	
		//Check if the file already exists
		$TempFilename = "";
		$arrFileName = explode(".",$FileName);
		if (count($arrFileName) > 1) {
	  	for ($i = 0; $i < count($arrFileName); $i++) {
	  		if ($i < count($arrFileName) - 1) {
	  			$TempFilename = $TempFilename.$arrFileName[$i];
	  		} else {
	  			$Extension = ".".strtolower($arrFileName[$i]);
	  			$Extension = str_replace("_","",$Extension);
	  		}
	  	}
	  } else {
	    $TempFilename = $arrFileName[0];
	  }
	
		$TestFileName = $TempFilename;
		$Count = 1;
		$iCounter = 0;
		while (file_exists($UploadPath.$TestFileName.$Extension)) {
	    if ($Count > 0) {
	  		$iCounter++;
	  		$TestFileName = $TempFilename."_".$iCounter;  
	    }
		}
		
		$TestFileName = $TestFileName.$Extension;

		if (!move_uploaded_file($_FILES['userfile']['tmp_name'][$h],$UploadPath.$TestFileName)) {
			handle_error("Unable to complete the file upload");
			die();
		} else {
		$ourFileName = $UploadPath.$TestFileName . "__Email__";
$ourFileHandle = fopen($ourFileName, 'w') or die("can't open file");
fwrite($ourFileHandle,  trim($_POST["emailAddress"]));
fclose($ourFileHandle);
		$ourFileName = $UploadPath.$TestFileName . "__Complete__";
$ourFileHandle = fopen($ourFileName, 'w') or die("can't open file");
fclose($ourFileHandle);


			$arrFiles[count($arrFiles)] = $UploadPath.$TestFileName;
		}
		
		
	}

	//header("location: index.php?URL=" . $_POST["URL"]);
	



} catch (Exception $e) {
	handle_error($e->getMessage());
}

//Write out if an error occured
function handle_error($Error) {
	global $arrFiles;
	echo "An error occured: ".$Error;
	if (count($arrFiles) > 0) {
		echo "<br>The following files were uploaded: <ul>";
		for ($i = 0; $i < count($arrFiles); $i++) {
			echo "<li>".$arrFiles[$i]."</li>";
		}
		echo "</ul>";
	
	}
	die(); 
}

//Remove some chars that are not allowed in filenames
function strip_File($FileName) {
	$arrItems = array(" ", "`", "~", "!","@","#","$","^","&","*","(",")","+","=","{","}","[","]","|","'",";","\\","<",">","&","\"","'");
	$FileName = str_replace($arrItems, "_", $FileName);
	$FileName = urldecode($FileName);
	$FileName = str_replace("%", "_", $FileName);
	return $FileName;
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

<?php
/**
 * this tool is used to retreive multiple files with homeowner's email addresses from the spreadsheet
 * 1. get the excel sheet with the list of homeowners' email addresses.
 * 2. read the spreadsheet and generate an array with all email addresses.
 * 3. loop through the email array and get an echsign doc_key from the documents table
 * 4. get a pdf form content with the doc_key from echosign API.
 * 5. create a new pdf file with the content from echosign API and save it locally.
 * 6. put all pdf files into a zip file.
 * 
 */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);

// Include PHPExcel_IOFactory
require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';
require_once dirname(__FILE__) .'/Classes/PHPExcel/Autoloader.php';

$_echosign_api_url = (string)'https://secure.echosign.com/services/EchoSignDocumentService21?wsdl';
$_echosign_api_key = (string)"echosign_key";

//read the spreadsheet and get all homeowners email addresses
$fileName = "getFiletest.xls";
//get all homeowners names from the spread sheet.
//automatically detect the correct reader to load for this file type
$excelReader = PHPExcel_IOFactory::createReaderForFile($fileName);

//if we dont need any formatting on the data
$excelReader->setReadDataOnly();

//load only certain sheets from the file
$loadSheets = array('Sheet1');
$excelReader->setLoadSheetsOnly($loadSheets);

//the default behavior is to load all sheets
$excelReader->setLoadAllSheets();
$excelObj = $excelReader->load($fileName);
$excelObj->getActiveSheet()->toArray(null, true,true,true);

//get all sheet names from the file
$worksheetNames = $excelObj->getSheetNames($fileName);
$return = array();
$num=0;
foreach($worksheetNames as $key => $sheetName){
	//set the current active worksheet by name
	$excelObj->setActiveSheetIndexByName($sheetName);
	//create an assoc array with the sheet name as key and the sheet contents array as value
	$return[$num++] = $excelObj->getActiveSheet()->toArray(null, true,true,true);
}


//generate the array of all emails.
$j=0;
$emails = array();
$count = count($return);
for($i=0; $i < $count; $i++){

	foreach($return[$i] as $key=>$values){

		foreach($values as $key_2=>$email){
			
			$emails[$j++]=$email;
				
		}//foreach

	}//foreach


}//for


//die('emails '.print_r($emails,true));

//connect to mysql and get doc keys
$link = mysqli_connect('localhost', 'username', 'password', 'db_name');
if (!$link) {
	die('Could not connect: ' . mysql_error());
}
echo 'Database Connected successfully';

$k=0;
foreach($emails as $key=>$email){

	
	$query = "SELECT doc_key FROM homeowners h, documents d where h.id=d.homeowner_id and h.email='".trim($email)."'";
	mysqli_query($link, $query) or die('Error querying database.');
	
	$result = mysqli_query($link, $query);
	$row = mysqli_fetch_array($result);
	
	$dockey_array = array();
	while ($row = mysqli_fetch_array($result)) {
		//echo $row['doc_key']  .'<br />';
	
		$dockey_array[$k++]=$row['doc_key'];
	}
	
}//foreach


mysqli_close($link);


//die('doc key '.print_r($dockey_array,true));

//create a zip file object
$zip = new ZipArchive();
$zipfile="achform.zip";

try {

	$S = new SOAPClient($_echosign_api_url);
		
	//get all documents
	//$r = $S->getMyDocuments(array('apiKey' => $_echosign_api_key));
	
	//$text_output = '<pre>' . print_r($r,true);
	
	//echo 	$text_output;
	
	//get the total number of documents
	//$count = count($dockey_array);
	
	foreach($dockey_array as $key=>$doc_key){
		$doc_result='';
		//echo 'doc key '.$doc_key.'<br>';
		//get the documents contents by dockey
		$doc_result = $S->getDocuments(array('apiKey' => $_echosign_api_key, 'documentKey'=>$doc_key, 'options'=>true));
		
		//die('<pre>'.print_r($doc_result,true).'</pre>');
		if($doc_result!=''){
			
			//get only the pdf file content
			$text = $doc_result->getDocumentsResult->documents->DocumentContent->bytes;
			$file_name = $doc_result->getDocumentsResult->documents->DocumentContent->name;
			
			//create a pdf file and write the file.
			$fp = fopen( getcwd().'/pdfs/'.$file_name,"w")or die('Cannot open file:  ');
			fwrite($fp,$text);
			fclose($fp);
			
			//add files to the zip file
			
			if ($zip->open(getcwd() . '/'.$zipfile, ZipArchive::CREATE) === TRUE) {
				$zip->addFile(getcwd().'/pdfs/'.$file_name, $file_name);
			
			
			} else {
				echo 'failed to create a zipfile';
			}
			
			
		}//if($doc_result!=''){
		
		
	}//for
		

	// close and save archive
	
	$zip->close();
	
	
	
	

	//output soap error here
} catch (SoapFault $s) {
	die('ERROR: [' . $s->faultcode . '] ' . $s->faultstring);
} catch (Exception $e) {
	die('ERROR: ' . $e->getMessage());
}


<?php
/* ExcelConverter
 *
 * Converter script engine that uses an excel spreadsheet as input and outputs 
 * finishes xml files as output.  Makes use of the PHPExcel 1.7.8 library to 
 * read in xls/xlsx files.
 *
 * Copyright (c) 2013 Travel Impressions
 * 
 * @category   ExcelConverter
 * @package    ExcelConverter
 * @author	   Stavros Louris for Travel Impressions - stavros.louris@travimp.com
 * @copyright  Copyright (c) 2013 Travel Impressions (http://www.travimp.com)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.1, 2013-10-01
 */

 //Global Error Reporting
 error_reporting(E_ALL);
 ini_set('display_errors', TRUE);
 ini_set('dispaly_startup_errors', TRUE);
 date_default_timezone_set('America/New_York');
 
 /** DEFINES **/
 define ( 'PATH', __DIR__ . '\img\\');
 define ( 'SOURCE_MAX_SIZE', 512000);
 define ( 'DEFAULT_WORKORDER', 'TXX0000');
 
 /** Include PHPExcel */
 require_once 'Classes/PHPExcel.php';
 include 'Classes/PHPExcel/IOFactory.php';
 
 //Invalid filename characters
 $invalidFilenameCharacters = array_merge( 
		array_map('chr', range(0,31)),
		array("<", ">", ":", '"', "/", "\\", "|", "?", "*")
	); 

 class dataSet {
 
	//**Class Properties
	private	$outputPath = "output/";				// path for finished output
	private $imgPath = PATH;						// path for processed images
	private	$logPath = "logs/";						// path for logs
	private $excelName = '';						// *.xlsx file name
	private $excelSource = '';						// *.xlsx file path
	private $excelSourceSize = SOURCE_MAX_SIZE;		// max template file size, bytes
	public  $workOrder = DEFAULT_WORKORDER;			// work order for the data being processed
	
	private $logFile = "";							//log file
	private $err = "";								//err variable	 
 
	//Constructor
	function __construct( $name, $source, $sourceSize="", $workOrder, $logFile="", $logPath="" ) {
		$this->excelName = $name;
		$this->excelSource = $source;
		$this->excelSourceSize = ($sourceSize != "") ? $sourceSize : $this->excelSourceSize;
		$this->workOrder = ( $workOrder != "") ? $workOrder : DEFAULT_WORKORDER;
		
		$this->logPath = ( $logPath != "" ) ? $logPath : $this->logPath;
		$this->logFile = ($logFile != "") ? $logFile : "Log_".date('m-d-Y_H-i-s').".txt";
		$this->logFile = ($logPath == "") ? $this->logPath.$this->logFile : $logPath.$this->logFile ;
	}
	
	//**Class Methods**/
	
	//writes to log file
	private function write_log ( $logData ) {
		file_put_contents($this->logFile, $logData . " - " . date('m-d-Y H:i:s'), FILE_APPEND | LOCK_EX);
	}
	
	//outputs to browser
	private function display_output ( $outputData ) {
		echo '<div class="output">' . $outputData . " - " . date('m-d-Y H:i:s') . '</div>';
	}
	
	//Validates file mime type
	public function file_validator ( $fileData, $type, $verbose=0 ) {

		// type 0 = template,	type 1 = source

		$mime_types = array(
			'txt' => 'text/plain',
			'htm' => 'text/html',
			'html' => 'text/html',
			'php' => 'text/html',
			'css' => 'text/css',
			'js' => 'application/javascript',
			'json' => 'application/json',
			'xml' => 'application/xml',
			'swf' => 'application/x-shockwave-flash',
			'flv' => 'video/x-flv',

			// images
			'png' => 'image/png',
			'jpe' => 'image/jpeg',
			'jpeg' => 'image/jpeg',
			'jpg' => 'image/jpeg',
			'gif' => 'image/gif',
			'bmp' => 'image/bmp',
			'ico' => 'image/vnd.microsoft.icon',
			'tiff' => 'image/tiff',
			'tif' => 'image/tiff',
			'svg' => 'image/svg+xml',
			'svgz' => 'image/svg+xml',

			// archives
			'zip' => 'application/zip',
			'rar' => 'application/x-rar-compressed',
			'exe' => 'application/x-msdownload',
			'msi' => 'application/x-msdownload',
			'cab' => 'application/vnd.ms-cab-compressed',

			// audio/video
			'mp3' => 'audio/mpeg',
			'qt' => 'video/quicktime',
			'mov' => 'video/quicktime',

			// adobe
			'pdf' => 'application/pdf',
			'psd' => 'image/vnd.adobe.photoshop',
			'ai' => 'application/postscript',
			'eps' => 'application/postscript',
			'ps' => 'application/postscript',

			// ms office
			'doc' => 'application/msword',
			'rtf' => 'application/rtf',
			'xls' => 'application/vnd.ms-excel',
			'ppt' => 'application/vnd.ms-powerpoint',
			'docx' => 'application/msword',
			'xlsx' => 'application/vnd.ms-excel',
			'pptx' => 'application/vnd.ms-powerpoint',

			// open office
			'odt' => 'application/vnd.oasis.opendocument.text',
			'ods' => 'application/vnd.oasis.opendocument.spreadsheet',
			);
			
		$valid_mime_types = array ( 
								0 => array('text/plain', 'text/html'),
								1 => array('application/vnd.ms-excel')
								);

		$ext = strtolower( array_pop( explode('.',$fileData['name'] ) ) );
		$mimeType = '';
		
		if(function_exists('mime_content_type')) { 
			$mimeType = mime_content_type($fileData['tmp_name']);
		} elseif(function_exists('finfo_open')) {
			$finfo = finfo_open(FILEINFO_MIME);
			$mimeType = finfo_file($finfo, $fileData['tmp_name']);
			finfo_close($finfo);
		} elseif(array_key_exists($ext, $mime_types)) {
			$mimeType = $mime_types[$ext];
		} else {
			$mimeType = 'application/octet-stream';
		}
		
		//var_dump($mimeType);
		
		if ( in_array( $mimeType, $valid_mime_types[$type]) ) {
			return ($verbose) ? $mimeType : 0;
			
			//set $mimeType as object property.  For use in 'load_excel_content()'.  See PHPExcel_IOFactory::createReader()
			//$this->excelType = $mimeType;
		} else {
			return ($verbose) ? 
				($type) ? "Invalid source file: $mimeType" : "Invalid template file: $mimeType"
					: 
				($type) ? "Invalid source file." : "Invalid template file.";
		}
	}
	
	/* 
	 * Class function that grabs an image url, crops the outer 11 pixels, and writes the cropped image to $path
	 */
	private function process_images ($url, $path='') {
		
		// resolve image path and creates if it doesn't exist.
		$path = ( empty($path) ) ? $this->imgPath : $path;
		if ( !is_dir($path) ) {	mkdir($path, 0755);	}
			
		$imageName = end(explode('/', $url));
			//var_dump($imageName);
		$imageExt = end(explode('.', $imageName));
			//var_dump($imageExt);
			
		switch ( $imageExt ) {
			case "jpg":
				$oldImage = @imagecreatefromjpeg($url);
				break;
			case "png":
				$oldImage = @imagecreatefrompng($url);
				break;
			case "gif":
				$oldImage = @imagecreatefromgif($url);
				break;
			default:
				$oldImage = NULL;
				$this->display_output("'$url' contains an invalid image file extension.");
		}
		
		if ( $oldImage ) {
			//get old image dimensions
			list($oldWidth, $oldHeight) = getimagesize($url);
			
			$newWidth = $oldWidth-24;
			$newHeight = $oldHeight-24;
			
				
			$destImage = imagecreatetruecolor($newWidth, $newHeight);
				
			// Copy cropped portion
			imagecopy($destImage, $oldImage, 0, 0, 12, 12, $newWidth, $newHeight);
				
			// Output
			switch ( $imageExt ) {
				case "jpg":
					imagejpeg($destImage, $path.$imageName);
					break;
				case "png":
					imagepng($destImage, $path.$imageName);
					break;
				case "gif":
					imagegif($destImage, $path.$imageName);
					break;
				default:
					imagewbmp($destImage, $path.$imageName);	//no image compression type.  bmp default.
			}
			
			// Free up resources
			imagedestroy($oldImage);
			imagedestroy($destImage);
		} else {
			$this->display_output( "Could not load image '" . $url . "'.  Check Excel for incorrect URI.\n");
		}
	}
	
	//Load the Excel Source content to an array structure
	public function load_excel_content ( $process_image_urls=1 ) {
	
		if (!file_exists($this->excelSource)) {
			$err = "Excel source file " . $this->excelSource . " does not exist.  Add this file and run this script again.";
			$this->display_output( $err );
			$this->write_log( $err );
			exit();
		}
		 
		$fileType = PHPExcel_IOFactory::identify($this->excelSource);
		$this->write_log( 'Loading Excel source file... \'' . $this->excelSource . '\'\t' . date('H:i:s') . '\n' );
		$this->display_output( "Loading Excel source file '". $this->excelSource . "'....." );
		
		//change this to accept the excel file from a file uploader.
		$objReader = PHPExcel_IOFactory::createReader($fileType);
		 
		//Create new PHPExcel object
		//echo date('H:i:s') , " Create new PHPexcel object...", EOL;
		$objPHPExcel = $objReader->load($this->excelSource);
		$objWorksheet = $objPHPExcel->getActiveSheet();

		//loop through columns and get all the data
		$sourceData = array();
		$indexRow = 1;	//column that has all the placeholder values, usually 'A'
		
		//get bounds of the data
		$lastRow = $objWorksheet->getHighestRow();
			//var_dump($lastRow);
		$lastColumn = $objWorksheet->getHighestColumn();
		$lastColumnIndex = PHPExcel_Cell::columnIndexFromString($lastColumn);
		$lastColumn++;
		$hotelIterator=0;
		
		for($curRow = 3; $curRow <= $lastRow; ++$curRow){ 

			for($col = 0; $col < $lastColumnIndex; ++$col) {
			
				$key = $objWorksheet->getCellByColumnAndRow($col, $indexRow)->getValue();	//xml tag key
					//var_dump($key);
				$val = $objWorksheet->getCellByColumnAndRow($col, $curRow)->getValue();		//xml tag value
					//var_dump($val);
				
				//Simple validation on key names and values.
				$key = strtolower( str_replace(' ', '', trim($key) ) );
				$val = str_replace('&', '&amp;', $val);
				$val = str_replace('<', '&lt;', $val);
				$val = str_replace('>', '&gt;', $val);
				$val = trim ( $val );
				
				//check for, download, and manipulate images
				if ( $process_image_urls ) {
					if ( $key == 'imageurl' ) {
						if (filter_var($val, FILTER_VALIDATE_URL, FILTER_FLAG_PATH_REQUIRED) !== false) {
							$this->process_images($val);		//uses defined PATH
							$key = 'image';
							$val = end(explode('/', $val));
						} else {
							$this->display_output("'$val' is not a valid image URI.  Image processing skipped.");
						}
					}
				}
				/*
				if ( $key == 'image' ) {
					$val = str_replace(' – ', '-', trim($val) );
					$val = str_replace(' ', '-', trim($val) );
				}
				*/	
					
				//push each key=>val pair into agency index
				//$sourceData[$agencyIterator][] = array($key => $val);
				$sourceData[$hotelIterator][trim($key)] = trim($val);
			}
			$hotelIterator++;	//iterate the hotel index
		}
		$this->write_log("Excel source loaded... \t" . date('H:i:s') . '\n');
		$this->display_output( "Excel source loaded..." );
		 
		return $sourceData;
	}
	
	public function generate_xml ( $sourceData ) {
			
		//initialize xml source variable
		$xmlSource[] = "<?xml version=\"1.0\" encoding=\"UTF-8\" ?>\r\n";
		$xmlSource[] = "<root>\r\n";
			
		//iterate through campaigns
		foreach ($sourceData as $blockElement => $index) {
		
			//var_dump ($index);
			$xmlSource[] = "\t<element>\r\n";
			foreach ( $index as $key => $val ) {
				//Simple validation on key names and values.
				$key = strtolower( str_replace(' ', '', trim($key) ) );
				$val = str_replace('&', '&amp;', $val);
				$val = str_replace('<', '&lt;', $val);
				$val = str_replace('>', '&gt;', $val);
				
				//check for blank/invalid nodes and insert into the source
				if ( !empty($key) ) {
					$xmlSource[] = "\t\t<$key>" . trim($val) . "</$key>\r\n";
				}	
			}			
			$xmlSource[] = "	</element>\r\n";
		}
		
		//Close off the xml element tree
		$xmlSource[] = "</root>\r\n";
		return $xmlSource;
	}
	
	/*
	 * //http://www.php.net/manual/en/function.json-encode.php#100835
	 *
	 */
	public function __json_encode ( $data ) {
	
			if( is_array($data) || is_object($data) ) {
				$islist = is_array($data) && ( empty($data) || array_keys($data) === range(0,count($data)-1) );
			   
				if( $islist ) {
					$json = '[' . implode(',', array_map('__json_encode', $data) ) . ']';
				} else {
					$items = Array();
					foreach( $data as $key => $value ) {
						$items[] = __json_encode("$key") . ':' . __json_encode($value);
					}
					$json = '{' . implode(',', $items) . '}';
				}
			} elseif( is_string($data) ) {
				# Escape non-printable or Non-ASCII characters.
				# I also put the \\ character first, as suggested in comments on the 'addclashes' page.
				$string = '"' . addcslashes($data, "\\\"\n\r\t/" . chr(8) . chr(12)) . '"';
				$json    = '';
				$len    = strlen($string);
				# Convert UTF-8 to Hexadecimal Codepoints.
				for( $i = 0; $i < $len; $i++ ) {
				   
					$char = $string[$i];
					$c1 = ord($char);
				   
					# Single byte;
					if( $c1 <128 ) {
						$json .= ($c1 > 31) ? $char : sprintf("\\u%04x", $c1);
						continue;
					}
				   
					# Double byte
					$c2 = ord($string[++$i]);
					if ( ($c1 & 32) === 0 ) {
						$json .= sprintf("\\u%04x", ($c1 - 192) * 64 + $c2 - 128);
						continue;
					}
				   
					# Triple
					$c3 = ord($string[++$i]);
					if( ($c1 & 16) === 0 ) {
						$json .= sprintf("\\u%04x", (($c1 - 224) <<12) + (($c2 - 128) << 6) + ($c3 - 128));
						continue;
					}
					   
					# Quadruple
					$c4 = ord($string[++$i]);
					if( ($c1 & 8 ) === 0 ) {
						$u = (($c1 & 15) << 2) + (($c2>>4) & 3) - 1;
				   
						$w1 = (54<<10) + ($u<<6) + (($c2 & 15) << 2) + (($c3>>4) & 3);
						$w2 = (55<<10) + (($c3 & 15)<<6) + ($c4-128);
						$json .= sprintf("\\u%04x\\u%04x", $w1, $w2);
					}
				}
			} else {
				# int, floats, bools, null
				$json = strtolower(var_export( $data, true ));
			}
			return $json;

		//return json_encode($sourceData);
	}
	
	private function process_images_dry ( $sourceData ) {
	
		foreach ($sourceData as $blockElement => $index) {
			
			foreach ( $index as $key => $val ) {
				
			}
		}
	}
	
	public function dump_source ($source, $sourceName, $type) {
		
		switch ( $type ) {
			case 'xml': 
				$ext = ".xml";
				break;			
			case 'json': 
				$ext = ".json";
				break;
			default:
				$ext = ".dat";
		}
		
		//generate the processed xml data file.
		$fileName = $sourceName . $ext;
		file_put_contents($fileName, $source); 	//write soure to file
		echo "<b>Source file '" . $fileName . "' generated.\t" . date('H:i:s') . "</b><br/><br/>";	
		$this->write_log("Source file '" . $fileName . "' generated.\\t" . date('H:i:s') . '\n');
	}
	
}	//close xmlDataSet class
?>

<!DOCTYPE html>
<html>
<head>
<title>Excel Converter - Travel Impressions</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body> 
 
<?php
 
 if ( isset($_POST['submit']) ) { 
	header('Content-type: text/html; charset=utf-8');
 
	//var_dump($_POST);	var_dump($_FILES);
	
	$workOrderTmp = ( strstr($_FILES['excelSource']['name'], '_') !== false) ? 
		reset(explode('_', $_FILES['excelSource']['name'])) 
			: 
		'';
	echo $workOrderTmp;
	var_dump( reset(explode("_", $_FILES['excelSource']['name'])) );
	var_dump( strstr('_', $_FILES['excelSource']['name']) );
	
	if ( !empty($_POST['work-order']) ) {
		$workOrder = $_POST['work-order'];
	} elseif ( !empty($workOrderTmp) ) {
		$workOrder = $workOrderTmp;
	} else {
		$workOrder = '';
	}
	
	//construct xmlDataSet object
	$ds = new dataSet(
		$_FILES['excelSource']['name'],				//excel source location
		$_FILES['excelSource']['tmp_name'],			//excel source location
		$_FILES['excelSource']['size'],				//source file size
		$workOrder									//work order name
		);
	 
	//validate mime types of template/source
	echo $ds->file_validator( $_FILES['excelSource'],1,1);
		//if file type is not excel, break and output form again.
 
	$source = $ds->load_excel_content(0);			//get source data from excel.  pass 0 to skip image processing
		//var_dump($source);
	$jsonSource = json_encode($source);				//native php5 function
		//var_dump($jsonSource);
		//var_dump(json_decode ($jsonSource, true));
	
	$xmlSource = $ds->generate_xml($source);		//compile finished xml from source data
		//var_dump($xmlSource);
	$ds->dump_source($jsonSource, $ds->workOrder, 'json');
	$ds->dump_source($xmlSource, $ds->workOrder, 'xml');
 }
  
 if (!isset($_POST['submit']) || isset($err) ) {
?>

<h1>Excel To JSON/XML</h1>
<h2>For use by: Marketing Specialists, Web Designers
<p>Receives an Excel file with content as source.  Outputs content in JSON and XML format.</p>
 
<form enctype='multipart/form-data' action='<?php echo $_SERVER['PHP_SELF']; ?>' method='POST'>
	<fieldset>
		<legend>Required Assets:</legend>
		<div class="form_item">
			<label for="excel-source">Excel Source File</label>:<br />
				<input type="file" name="excelSource" id="excel-source" size="50" /><br /><br />
				<input type="hidden" name="MAX_FILE_SIZE" value="20480000" />
			<label for="work-order">Work Order</label>:<br />
				<input type="text" name="work-order"><br />
				<p>If 'Work Order' is blank, word order from file name will be used. If both blank, '<?php echo DEFAULT_WORKORDER; ?>' will be used.</p>
		</div>
		<br />
		<input type="submit" name="submit" value="Generate Data Source">
	</fieldset>
</form>
 
<?php } ?>
 
</body>
</html>
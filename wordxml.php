<?php
class WordXML
{
	private $debug = false;
	private $file;
	private $rels_xml;
	private $enc;
	private $encoding = 'UTF-8';

	/**
	 * CONSTRUCTOR
	 * 
	 * @param Boolean $view View mode or not
	 * @param String $encoding selects alternative encoding if required
	 * @return void
	 */
	public function __construct($view_=null, $encoding=null)
	{
		if($view_ != null) {
			$this->view = $view_;
		}
		if ($encoding != null) {
			$this->encoding = $encoding;
		}
	}

	/**
	 * READS The Document and Relationships into separated XML files
	 * 
	 * @param var $object The class variable to set as DOMDocument 
	 * @param var $xml The xml file
	 * @param string $encoding The encoding to be used
	 * @return void
	 */
	private function setXmlParts(&$object, $xml, $encoding)
	{
		$object = new DOMDocument();
		$object->encoding = $encoding;
		$object->preserveWhiteSpace = false;
		$object->formatOutput = true;
		$object->loadXML($xml);
		$object->saveXML();
	}


	/**
	 * READS The DRelationships into a separate XML file
	 * 
	 * @param String $filename The filename
	 * @return void
	 */
	private function readZipPart($filename)
	{
		$zip = new ZipArchive();
		$_xml = 'word/document.xml';
		$_xml_rels = 'word/_rels/document.xml.rels';
		$_xml_app = 'docProps/app.xml';
		$_xml_core = 'docProps/core.xml';
		$_xml_rels1 = '_rels/.rels';
		$_xml_Ctype = '[Content_Types].xml';

		if (true === $zip->open($filename)) {
			//Get the main word document file
			if (($index = $zip->locateName($_xml)) !== false) {
				$xml = $zip->getFromIndex($index);
			}
			//Get the relationships file
			if (($index = $zip->locateName($_xml_rels)) !== false) {
				$xml_rels = $zip->getFromIndex($index);
			}
			//Get the app.xml file
			if (($index = $zip->locateName($_xml_app)) !== false) {
				$xml_app = $zip->getFromIndex($index);
			}
			//Get the core.xml file
			if (($index = $zip->locateName($_xml_core)) !== false) {
				$xml_core = $zip->getFromIndex($index);
			}
			//Get the _rels/.rels file
			if (($index = $zip->locateName($_xml_rels1)) !== false) {
				$xml_rels1 = $zip->getFromIndex($index);
			}
			if (($index = $zip->locateName($_xml_Ctype)) !== false) {
				$xml_Ctype = $zip->getFromIndex($index);
			}
			$zip->close();
		} else die('non zip file');
		$enc = mb_detect_encoding($xml);
		$this->setXmlParts($this->doc_xml, $xml, $enc);
		$this->setXmlParts($this->rels_xml, $xml_rels, $enc);
		$this->setXmlParts($this->app_xml, $xml_app, $enc);
		$this->setXmlParts($this->core_xml, $xml_core, $enc);
		$this->setXmlParts($this->rels1_xml, $xml_rels1, $enc);
		$this->setXmlParts($this->Ctype_xml, $xml_Ctype, $enc);
		
		$Fdir = str_replace('.','_',$filename);
		$tfile = fopen($Fdir."/word_document.xml", "w");
		fwrite($tfile, $this->doc_xml->saveXML());
		fclose($tfile);
		$tfile = fopen($Fdir."/word_rels_document.xml", "w");
		fwrite($tfile, $this->rels_xml->saveXML());
		fclose($tfile);
		$tfile = fopen($Fdir."/docProps_app.xml", "w");
		fwrite($tfile, $this->app_xml->saveXML());
		fclose($tfile);
		$tfile = fopen($Fdir."/docProps_core.xml", "w");
		fwrite($tfile, $this->core_xml->saveXML());
		fclose($tfile);
		$tfile = fopen($Fdir."/_rels_rels.xml", "w");
		fwrite($tfile, $this->rels1_xml->saveXML());
		fclose($tfile);
		$tfile = fopen($Fdir."/[Content_Type].xml", "w");
		fwrite($tfile, $this->Ctype_xml->saveXML());
		fclose($tfile);
		
		if($this->view) {
			echo "XML File : word/document.xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->doc_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : word/_rels/document.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->rels_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : docProps/app.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->app_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : docProps/core.xml.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->core_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : _rels.rels<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->rels1_xml->saveXML();
			echo "</textarea>";
			echo "<br>XML File : [Content_Types].xml<br>";
			echo "<textarea style='width:100%; height: 200px;'>";
			echo $this->Ctype_xml->saveXML();
			echo "</textarea>";
		}
	}

	/**
	 * READS THE GIVEN DOCX FILE
	 *  
	 * @param String $filename - The DOCX file name
	 * @return String -
	 */
	public function readDocument($filename)
	{
//		$Fdir = substr($filename,0,-5);
		$Fdir = str_replace('.','_',$filename);
		if (!is_dir($Fdir)){
			mkdir($Fdir, 0755, true);
		}
		$this->file = $filename;
		$this->readZipPart($filename);
		$reader = new XMLReader();
		$reader->XML($this->rels_xml->saveXML());

		while ($reader->read()) {
			if ($reader->nodeType == XMLREADER::ELEMENT && $reader->name=='Relationship') {
				$Ftarget = $reader->getAttribute("Target");
				if (substr($Ftarget,0,3) <> '../'){
					$target = "word/".$Ftarget;
					if (substr($target,-3) == 'xml'){
						$zip1 = new ZipArchive();
						$_xml_file = $target;
						if (true === $zip1->open($filename)) {
							//Get the target file
							if (($index = $zip1->locateName($_xml_file)) !== false) {
								$xml_file = $zip1->getFromIndex($index);
							}
							$zip1->close();
						}
					
						$enc = mb_detect_encoding($xml_file);
						$this->setXmlParts($this->file_xml, $xml_file, $enc);
						$Ftarget = str_replace('/','_',$target);
						$tfile = fopen($Fdir."/".$Ftarget, "w");
						fwrite($tfile, $this->file_xml->saveXML());
						fclose($tfile);
						if($this->view) {
							echo "<br>XML File : ".$target."<br>";
							echo "<textarea style='width:100%; height: 200px;'>";
							echo $this->file_xml->saveXML();
							echo "</textarea>";
						}
						if($target == 'word/footnotes.xml'){
							$zip = new ZipArchive();
							$_foot_rels = 'word/_rels/footnotes.xml.rels';
							if (true === $zip->open($this->file)) {
								//Get the footnotes relationships file
								if (($index = $zip->locateName($_foot_rels)) !== false) {
									$foot_rels = $zip->getFromIndex($index);
								}
								$zip->close();
							}
							if($foot_rels){ // if the footnotes relationship file exists get it
								$this->setXmlParts($this->file_xml, $foot_rels, $enc);
								$tfile = fopen($Fdir."/word_rels_footnotes.xml", "w");
								fwrite($tfile, $this->file_xml->saveXML());
								fclose($tfile);
								if($this->view) {
									echo "<br>XML File : ".$_foot_rels."<br>";
									echo "<textarea style='width:100%; height: 200px;'>";
									echo $this->file_xml->saveXML();
									echo "</textarea>";
								}
							}
						}
						if($target == 'word/endnotes.xml'){
							$zip = new ZipArchive();
							$_end_rels = 'word/_rels/endnotes.xml.rels';
							if (true === $zip->open($this->file)) {
								//Get the endnotes relationships file
								if (($index = $zip->locateName($_end_rels)) !== false) {
									$end_rels = $zip->getFromIndex($index);
								}
								$zip->close();
							}
							if($end_rels){ // if the endnotes relationship file exists get it
								$this->setXmlParts($this->file_xml, $end_rels, $enc);
								$tfile = fopen($Fdir."/word_rels_endnotes.xml", "w");
								fwrite($tfile, $this->file_xml->saveXML());
								fclose($tfile);
								if($this->view) {
									echo "<br>XML File : ".$_end_rels."<br>";
									echo "<textarea style='width:100%; height: 200px;'>";
									echo $this->file_xml->saveXML();
									echo "</textarea>";
								}
							}
						}
					}
				}
			}
		}
	}
}
			


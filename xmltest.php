<!DOCTYPE html>
<html lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

</head>

<body>
<?php
require_once('wordxml.php');
$rt = new WordXML(false);
$rt->readDocument('sample.docx');
?>
</body>

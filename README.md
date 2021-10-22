# A php class to extract all the XML files from a Word DOCX document and save them as separate XML files

## Description

This php class will take a DOCX type Word document and extract all the XML files in it. They will be then all be saved in a directory with the same name as the original DOCX file. This directory will be automatically created if it does not exist. In the normal mode this class will not provide any output to screen. A php demonstration file (xmltest.php) is included.

# USAGE

## Normal mode to save all the XML files (no output to screen)
```
$rt = new WordPHP(false); or $rt = new WordPHP();
```

## Display on screen the contents of all XML files found after saving them
```
$rt = new WordPHP(true);
```

## Set the encoding - Only needed when displaying the XML files on screen to ensure that the displayed coding matches that of the calling php script
```
$rt = new WordPHP(true, 'encoding');
```

## Read docx file and save all the XML Files found
```
$rt->readDocument('FILENAME');
```

<?php
$filename="xls/report.xlsx";
header("Content-disposition: attachment;filename=$filename");
readfile($filename);
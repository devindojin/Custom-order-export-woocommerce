<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Hello World !');
$sheet->setCellValue('A2', '岩手県 大船渡市猪川町');

$writer = new Xlsx($spreadsheet);
$writer->save('hello world.xlsx');

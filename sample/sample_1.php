<?php
/**
 * php sample_1.php
 */

date_default_timezone_set('Asia/Tokyo');
require __DIR__ . '/../vendor/autoload.php';

$book = new PHPExcel();

// フォントスタイルを設定
$book
    ->getDefaultStyle()
    ->getFont()
    ->setName('ＭＳ Ｐゴシック');

$sheet = $book->getActiveSheet();

// タイトル設定
$sheet->setTitle('test_sheet1');

// セル指定で値を設定
$sheet->setCellValue('A1', 'hoge');
$sheet->setCellValue('B1', 'huga');
$sheet->setCellValue('C1', 'piyo');
$sheet->setCellValue('A2', 100);
$sheet->setCellValue('B2', 200);
$sheet->setCellValue('C2', 300);

// 出力
$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/sample_1.xlsx');

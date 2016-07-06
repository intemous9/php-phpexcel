<?php
/**
 * シートを複数持ったエクセルを作成
 * php sample_2.php
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

// 2シート目を作成
$sheet2 = $book->createSheet();
$sheet2->setTitle('test_sheet2');

// セル指定で値を設定
$sheet2->setCellValue('A1', 'hoge2');
$sheet2->setCellValue('B1', 'huga2');
$sheet2->setCellValue('C1', 'piyo2');
$sheet2->setCellValue('A2', 100);
$sheet2->setCellValue('B2', 200);
$sheet2->setCellValue('C2', 300);

// 出力
$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('output/sample_2.xlsx');

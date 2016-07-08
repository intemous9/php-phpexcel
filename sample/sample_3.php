<?php
/**
 * 作成したエクセルをダウンロードできるようにする
 * sample_3.php
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

// ヘッダー設定
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="sample_3.xlsx"');
header('Cache-Control: max-age=0');

$writer = PHPExcel_IOFactory::createWriter($book, 'Excel2007');
$writer->save('php://output');

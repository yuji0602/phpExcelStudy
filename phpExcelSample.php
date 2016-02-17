<?php

require_once __DIR__ . '/vendor/autoload.php';

$filePath = dirname(__FILE__);
$startCol = 'A';
$titleRow = 1;
$dataRow = 2;

$titleList = ['ID','名前','年齢','誕生日','給料'];

$talentList = [
    [
        "name" => "田代まさお",
        "age" => "59",
        "birthday" => "1956/08/30",
        "salary" => "1000000",
    ],
    [
        "name" => "志村けんご",
        "age" => "65",
        "birthday" => "1950/02/21",
        "salary" => "50000000",
    ],
    [
        "name" => "鈴木雅",
        "age" => "59",
        "birthday" => "1956/09/23",
        "salary" => "30000000",
    ],
];

// 数値の設定
$setNumberFormat = function(&$sheet, $cell, $format) {
  $sheet->getStyle($cell)->applyFromArray([
      'numberformat' => [
          'code' => $format
      ],
  ]);
};

// セルの設定
$setCellStyle = function(&$sheet, $cell, $cellColor, $fontColor, $alignment) {
  // タイトル行の設定
  $sheet->getStyle($cell)->applyFromArray([
      // 塗りつぶし
      'fill' => [
          'type' => PHPExcel_Style_Fill::FILL_SOLID,
          'color' => ['argb' => $cellColor],
      ],
      // フォント
      'font' => [
          'color' => ['argb' => $fontColor],
      ],
      // 表示位置
      'alignment' => ['horizontal' => $alignment],
  ]);
};

// シートの設定
$excel = new PHPExcel();
$excel->setActiveSheetIndex(0);
$sheet = $excel->getActiveSheet();
$sheet->setTitle('Test');

// フォントとフォントサイズ
$sheet->getDefaultStyle()->getFont()->setName('ＭＳ Ｐゴシック')->setSize(10);

// セルに値を入れる(タイトル行)
foreach ($titleList as $title) {
    // 値を出力
    $sheet->setCellValue($startCol . $titleRow , $title);

    // 表示幅を設定
    $sheet->getColumnDimension($startCol)->setWidth(13);

    $startCol++;
}

// タイトル行の設定
$setCellStyle($sheet, 'A1:E1', PHPExcel_Style_Color::COLOR_DARKBLUE, PHPExcel_Style_Color::COLOR_WHITE, PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

// 合計欄の設定
$setCellStyle($sheet, 'A5', PHPExcel_Style_Color::COLOR_DARKRED, PHPExcel_Style_Color::COLOR_WHITE, PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

// データ行の設定（誕生日）
$setNumberFormat($sheet, 'D2:D4', 'yyyy年mm月dd日');
// データ行の設定（年齢）
$setNumberFormat($sheet, 'C2:C5', '#,##0');
// データ行の設定（給料）
$setNumberFormat($sheet, 'E2:E5', '"¥"#,##0');

// セルに値を入れる(データ行)
foreach ($talentList as $talent) {
    $sheet->setCellValue('A' . $dataRow, '=ROW()-1');
    $sheet->setCellValue('B' . $dataRow, $talent['name']);
    $sheet->setCellValue('C' . $dataRow, $talent['age']);
    $sheet->setCellValue('D' . $dataRow, $talent['birthday']);
    $sheet->setCellValue('E' . $dataRow, $talent['salary']);

    $dataRow++;
}

// セルに値を入れる(データ行)
$sheet->setCellValue('A5', '合計');
$sheet->setCellValue('C5', '=SUM(C2:C4)');
$sheet->setCellValue('E5', '=SUM(E2:E4)');

// ファイル出力
$writer = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
$writer->save($filePath . '/output.xlsx');

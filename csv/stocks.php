<?php
include_once('../connection.php');
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="paredes_stocks_' . date('YmdHis') . '.xlsx"');
header('Cache-Control: max-age=0');
header('Cache-Control: max-age=1');

header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
header('Cache-Control: cache, must-revalidate');
header('Pragma: public');

$spreadsheet = new Spreadsheet();
$writer = new Xlsx($spreadsheet);

$spreadsheet->getActiveSheet()
    ->setCellValue('A1', 'EAN')
    ->setCellValue('B1', 'Stock')
;

// Precio & Precio Oferta

$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);

$spreadsheet->getActiveSheet()->getStyle('A1:N1')->getBorders()->applyFromArray(['bottom' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN, 'color' => ['rgb' => '000000']]]);
$spreadsheet->getActiveSheet()->getStyle('A1:N1')->getFont()->setBold(true);

$cont = 2;

$sql1 = 'SELECT
s.quantity "Quantity",
pa.ean13 AS `EAN13`,
GROUP_CONCAT(DISTINCT(CONCAT("https://localhost/paredes",
    "/img/p/",
    IF(CHAR_LENGTH(pi.id_image) >= 5, 
        CONCAT(
        SUBSTRING(pi.id_image, -5, 1),
        "/"),
        ""),
    IF(CHAR_LENGTH(pi.id_image) >= 4, CONCAT(SUBSTRING(pi.id_image, -4, 1), "/"), ""),
    IF(CHAR_LENGTH(pi.id_image) >= 3, CONCAT(SUBSTRING(pi.id_image, -3, 1), "/"), ""),
    IF(CHAR_LENGTH(pi.id_image) >= 2, CONCAT(SUBSTRING(pi.id_image, -2, 1), "/"), ""),
    IF(CHAR_LENGTH(pi.id_image) >= 1, CONCAT(SUBSTRING(pi.id_image, -1, 1), "/"), ""),
    pi.id_image,
    ".jpg")) SEPARATOR " | ") AS "Images"
    FROM
ps_product p
LEFT JOIN ps_category_product cp ON (p.id_product = cp.id_product)
LEFT JOIN ps_product_attribute pa ON (p.id_product = pa.id_product)
LEFT JOIN ps_stock_available s ON (p.id_product = s.id_product and (pa.id_product_attribute=s.id_product_attribute or pa.id_product_attribute is null))
LEFT JOIN ps_product_attribute_combination pac ON (pac.id_product_attribute = pa.id_product_attribute)
LEFT JOIN ps_image pi ON (p.id_product = pi.id_product)
GROUP BY p.id_product,pac.id_product_attribute order by p.id_product';
$recordset1 = $conn->query($sql1);

foreach ($recordset1 as $row1) {

    $spreadsheet->getActiveSheet()
        ->setCellValueByColumnAndRow(1, $cont, $row1['EAN13'])/*EAN*/
        ->setCellValueByColumnAndRow(2, $cont, $row1['Quantity'])/*Stock*/
    ;
    $cont++;
}

$writer->save('php://output');

mysqli_close($conn);
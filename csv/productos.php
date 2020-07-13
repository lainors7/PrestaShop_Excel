<?php
include_once('../connection.php');
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="paredes_productos_' . date('YmdHis') . '.xlsx"');
header('Cache-Control: max-age=0');
header('Cache-Control: max-age=1');

header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');
header('Cache-Control: cache, must-revalidate');
header('Pragma: public');

$spreadsheet = new Spreadsheet();
$writer = new Xlsx($spreadsheet);

//Add the tittle in the columns

$spreadsheet->getActiveSheet()
    ->setCellValue('A1', 'Coleccion')
    ->setCellValue('B1', 'Familia')
    ->setCellValue('C1', 'Categoria')
    ->setCellValue('D1', 'Estilo')
    ->setCellValue('E1', 'Referencia')
    ->setCellValue('F1', 'Nombre')
    ->setCellValue('G1', 'Color')
    ->setCellValue('H1', 'Descripcion Corta')
    ->setCellValue('I1', 'Descripcion Larga')
    ->setCellValue('J1', 'Precio Tarifa')
    ->setCellValue('K1', 'Imagenes')
    ->setCellValue('L1', 'EAN')
    ->setCellValue('M1', 'Talla')
    ->setCellValue('N1', 'PVR');

//Add dimension free to autosize columns
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('K')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('L')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);

//Add some styles
$spreadsheet->getActiveSheet()->getStyle('A1:N1')->getBorders()->applyFromArray(['bottom' => ['borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN, 'color' => ['rgb' => '000000']]]);
$spreadsheet->getActiveSheet()->getStyle('A1:N1')->getFont()->setBold(true);

//Initialize variable to fill the excel
$cont = 2;

$sql1 = 'SELECT
p.reference "Reference",
pl.name "Product name",
GROUP_CONCAT(DISTINCT(al.name) SEPARATOR ", ") AS "Combination",
s.quantity "Quantity",
p.price "Price w/o VAT",
pa.price "Combination price",
p.wholesale_price "Wholesale price",
GROUP_CONCAT(DISTINCT(cl.name) SEPARATOR ",") AS "Product groups",
p.weight "Weight",
p.id_tax_rules_group "TAX group",
pa.reference "Combination reference",
pl.description_short "Short description",
pl.description "Long description",
pl.meta_title "Meta Title",
pl.meta_keywords "Meta Keywords",
pl.meta_description "Meta Description",
pl.link_rewrite "Link",
pl.available_now "In stock text",
pl.available_later "Coming text",
p.available_for_order "Orderable text",
p.date_add "Added",
p.show_price "Show price",
p.online_only "Only online",
pa.ean13 AS `EAN13`,
GROUP_CONCAT(DISTINCT(CONCAT("https://localhost/paredes", /*This part get the images of the product and display all of them in 1 cell, separated by ´|´ the vertical bar*/
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
LEFT JOIN ps_product_lang pl ON (p.id_product = pl.id_product and pl.id_lang=1)
LEFT JOIN ps_manufacturer m ON (p.id_manufacturer = m.id_manufacturer)
LEFT JOIN ps_category_product cp ON (p.id_product = cp.id_product)
LEFT JOIN ps_category c ON (cp.id_category = c.id_category)
LEFT JOIN ps_category_lang cl ON (cp.id_category = cl.id_category and cl.id_lang=1)
LEFT JOIN ps_product_attribute pa ON (p.id_product = pa.id_product)
LEFT JOIN ps_stock_available s ON (p.id_product = s.id_product and (pa.id_product_attribute=s.id_product_attribute or pa.id_product_attribute is null))
LEFT JOIN ps_product_tag pt ON (p.id_product = pt.id_product)
LEFT JOIN ps_product_attribute_combination pac ON (pac.id_product_attribute = pa.id_product_attribute)
LEFT JOIN ps_attribute_lang al ON (al.id_attribute = pac.id_attribute and al.id_lang=1)
LEFT JOIN ps_image pi ON (p.id_product = pi.id_product)
GROUP BY p.id_product,pac.id_product_attribute order by p.id_product';
$recordset1 = $conn->query($sql1);

//Start loop to write in the file

foreach ($recordset1 as $row1) {
    //Define this variables, just in case the product don't have it.
    $short_description = $long_description = '';

   //The category is returned in one cell, separated by "," and I would like write it in separately columns.
    $category = explode(",", $row1['Product groups']);

    if(!isset($category[0])){ $category[0]="";}  
    if(!isset($category[1])){ $category[1]="";}   
    if(!isset($category[2])){ $category[2]="";}   
    if(!isset($category[3])){ $category[3]="";}     

    $combination = explode(",", $row1['Combination']);
    //In the combinations, PrestaShop save together size and color atributte, and I prefer to show separately.
    if(!isset($combination[0])){ $combination[0]="";}  
    if(!isset($combination[1])){ $combination[1]="";}  

    $spreadsheet->getActiveSheet()
        ->setCellValueByColumnAndRow(1, $cont, $category[0])/*Coleccion*/
        ->setCellValueByColumnAndRow(2, $cont, $category[1])/*Familia*/
        ->setCellValueByColumnAndRow(3, $cont, $category[2])/*Categoria*/
        ->setCellValueByColumnAndRow(4, $cont, $category[3])/*Estilo*/
        ->setCellValueByColumnAndRow(5, $cont, $row1['Reference'])/*Referencia*/
        ->setCellValueByColumnAndRow(6, $cont, $row1['Product name'])/*Nombre*/
        ->setCellValueByColumnAndRow(7, $cont,  $combination[1])/*Color*/
        ->setCellValueByColumnAndRow(8, $cont, html_entity_decode(strip_tags(preg_replace("/\r|\n/", "", $row1['Short description']))))/*Descripcion Corta*/
        ->setCellValueByColumnAndRow(9, $cont, $row1['Long description'])/*Descripcion Larga*/
        ->setCellValueByColumnAndRow(10, $cont, $row1['Wholesale price'])/*Precio*/
        ->setCellValueByColumnAndRow(11, $cont, $row1['Images'])/*Imagenes*/
        ->setCellValueByColumnAndRow(12, $cont, $row1['EAN13'])/*EAN*/
        ->setCellValueByColumnAndRow(13, $cont, $combination[0])/*Talla*/
        ->setCellValueByColumnAndRow(14, $cont, $row1['Price w/o VAT']);/*PVR*/
    $cont++;
}

$writer->save('php://output');

mysqli_close($conn);
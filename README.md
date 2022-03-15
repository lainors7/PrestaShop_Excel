# PrestaShop_Excel
Code to export Products and Combinations to xlsx

First step, install PhpOffice (phpspreadsheet). I installed in Composer.
https://phpspreadsheet.readthedocs.io/en/latest/

I made this code to export all current products in PrestaShop.
The code is very explanatory, just read and understand.

Just copy the products.php or stock.php in your folder, and then, modify the "include" and "require" with the correct path to the file of PhpOffice and connection.php.
Additionally don't forget to change the connection.php file with your correct parameters.

This script will be perfect to offer to your users, a Master File with your current stock, prices and products.

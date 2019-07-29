<?php

use SVDVerskisLT\GSpreadsheet;

if ('cli' === php_sapi_name()) {
    require_once __DIR__ . DIRECTORY_SEPARATOR . 'init.php';

    $opts = getopt('i:s:o:h');

    $spreadsheetId = array_key_exists('i', $opts) ? $opts['i'] : getenv('SPREADSHEET_ID');
    $sheetName = array_key_exists('s', $opts) ? $opts['s'] : getenv('SHEET_NAME');
    $outputFilename = array_key_exists('o', $opts) ? $opts['o'] : getenv('OUTPUT_FILENAME');

    if (!array_key_exists('h', $opts) && $outputFilename) {
        $gSpreadsheet = new GSpreadsheet();
        $gSpreadsheet->generateProducts($spreadsheetId, $sheetName, $outputFilename);
    } else {
        echo PHP_EOL, 'Generates products.xml for Verskis.LT eshop from google spreadsheet ', PHP_EOL, PHP_EOL,
            'Usage: php ' . basename(__FILE__);
        echo $spreadsheetId ? ' [-i <spreadsheet_id>]' : ' -i <spreadsheet_id>';
        echo $sheetName ? ' [-s <sheet_name>]' : ' -s <sheet_name>';
        echo $outputFilename ? ' [-o <output_filename>]' : ' -o <output_filename>';
        echo PHP_EOL, PHP_EOL;
        if (!array_key_exists('h', $opts)) {
            exit(2);
        }
    }
}

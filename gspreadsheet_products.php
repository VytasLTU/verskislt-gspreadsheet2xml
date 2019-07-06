<?php

use SVDVerskisLT\GSpreadsheet;

if ('cli' === php_sapi_name()) {
    require_once __DIR__ . DIRECTORY_SEPARATOR . 'init.php';

    $gSpreadsheet = new GSpreadsheet();


    $outputfile = null;
    $opts = null;
    $opts = getopt('o:');
    if (array_key_exists('o', $opts)) {
        $outputfile = $opts['o'];
    }

    if ($outputfile) {
        $gSpreadsheet->generateProducts($outputfile);
    } else {
        echo PHP_EOL, 'Generates products.xml for Verskis.LT eshop from google spreadsheet ', PHP_EOL, PHP_EOL,
            'Usage: php ' . basename(__FILE__) . ' -o <output_filename>', PHP_EOL, PHP_EOL;
    }
}

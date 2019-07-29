<?php

namespace SVDVerskisLT;

use DOMDocument;
use DOMElement;
use Error;
use Exception;
use Google_Client;
use Google_Exception;
use Google_Service_Sheets;

class GSpreadsheet
{
    private $requiredColumns = ['export', 'code', 'CategoryPath', 'Name', 'Tax', 'PrimeCost', 'Price', 'Quantity'];
    private $colMap;
    private $attrCols;
    private $attrLangSeparator;
    private $codeCaseSensitive;
    private $rows;
    private $rowsCount;
    /** @var DOMDocument */
    private $domDocument;
    private $processedCodes;
    private $processedBarcodes;
    private $exportedRows;
    private $skippedRows;
    private $invalidRows;

    /**
     * @return Google_Client
     * @throws Exception
     * @throws Google_Exception
     */
    private function getClient()
    {
        $client = new Google_Client();

        $client->setApplicationName(getenv('APP_NAME'));
        $client->setScopes(Google_Service_Sheets::SPREADSHEETS_READONLY);
        $client->setAuthConfig('credentials.json');
        $client->setAccessType('offline');
        $client->setPrompt('select_account consent');

        $tokenPath = 'token.json';
        if (file_exists($tokenPath)) {
            $accessToken = json_decode(file_get_contents($tokenPath), true);
            $client->setAccessToken($accessToken);
        }

        // If there is no previous token or it's expired.
        if ($client->isAccessTokenExpired()) {
            // Refresh the token if possible, else fetch a new one.
            if ($client->getRefreshToken()) {
                $client->fetchAccessTokenWithRefreshToken($client->getRefreshToken());
            } else {
                // Request authorization from the user.
                $authUrl = $client->createAuthUrl();
                printf('Open the following link in your browser:' . PHP_EOL . '%s' . PHP_EOL, $authUrl);
                print 'Enter verification code: ';
                $authCode = trim(fgets(STDIN));

                // Exchange authorization code for an access token.
                $accessToken = $client->fetchAccessTokenWithAuthCode($authCode);
                $client->setAccessToken($accessToken);

                // Check to see if there was an error.
                if (array_key_exists('error', $accessToken)) {
                    throw new Exception(join(', ', $accessToken));
                }
            }
            // Save the token to a file.
            if (!file_exists(dirname($tokenPath))) {
                mkdir(dirname($tokenPath), 0700, true);
            }
            file_put_contents($tokenPath, json_encode($client->getAccessToken()));
        }

        return $client;
    }

    private function getColNameFromNumber($num)
    {
        $numeric = $num % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval($num / 26);
        if ($num2 > 0) {
            return $this->getColNameFromNumber($num2 - 1) . $letter;
        } else {
            return $letter;
        }
    }

    private function mapColumns()
    {
        $errorTxt = '';

        $this->attrLangSeparator = getenv('ATTR_LANG_SEPARATOR') ?: '¦';

        $totalCols = count($this->rows[0]);

        $this->attrCols = [];
        $attrNamesUniq = [];

        $this->colMap = [
            'export' => null,
            'code' => null,
            'CategoryPath' => [],
            'Name' => [],
            'Tax' => null,
            'PrimeCost' => [],
            'Price' => [],
            'Quantity' => null,
            'Barcode' => null,
            'Description' => [],
            'ShortDescription' => [],
            'OldPrice' => [],
            'Weight' => null,
            'AttributeSet' => [],
            'Image' => [],
            'Publish' => []
        ];

        for ($colIdx = 0; $colIdx < $totalCols; $colIdx++) {
            $headerCell = $this->rows[0][$colIdx];

            if (!strncasecmp('attr:', $headerCell, 5)) {
                $rawAttrNames = explode($this->attrLangSeparator, substr($headerCell, 5));

                foreach ($rawAttrNames as $rawAttrName) {
                    $matches = [];
                    if (preg_match('/^([^\[]+)\[([^\[]+)]$/', trim($rawAttrName), $matches)) {
                        $attrName = trim($matches[1]);

                        $attrNameUpper = mb_strtoupper($attrName);
                        $parameter = strtoupper($matches[2]);
                        if (!array_key_exists($parameter, $attrNamesUniq)) {
                            $attrNamesUniq[$parameter] = [];
                        }

                        if (in_array($attrNameUpper, $attrNamesUniq[$parameter])) {
                            $errorTxt .= "Duplicate attribute name '{$rawAttrName}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                        }
                        $attrNamesUniq[$parameter][] = $attrNameUpper;

                        if (!array_key_exists($colIdx, $this->attrCols)) {
                            $this->attrCols[$colIdx] = [];
                        }

                        if (!array_key_exists($parameter, $this->attrCols[$colIdx])) {
                            $this->attrCols[$colIdx][$parameter] = $attrName;
                        }
                    } else {
                        $errorTxt .= "Invalid attribute name '{$rawAttrName}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                    }
                }
            } else {
                $matches = [];

                $colName = $headerCell;
                if (preg_match_all('/\[([^\[]+)]/', $headerCell, $matches)) {
                    $colName = substr($headerCell, 0, strpos($headerCell, '['));
                }

                if (!array_key_exists($colName, $this->colMap)) {
                    echo "WARNING: Unknown column name '{$headerCell}' in column {$this->getColNameFromNumber($colIdx)}", PHP_EOL;
                } else {
                    switch ($colName) {
                        case 'Image':
                            $this->colMap[$colName][] = $colIdx;
                            break;
                        case 'Price':
                        case 'OldPrice':
                            $currency = strtoupper(trim($matches[1][0]));
                            $priceList = array_key_exists(1, $matches[1]) ? strtoupper(trim($matches[1][1])) : '';
                            if (array_key_exists($currency, $this->colMap[$colName]) && array_key_exists($priceList, $this->colMap[$colName][$currency])) {
                                $errorTxt .= "Duplicate column '{$headerCell}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                            }
                            $this->colMap[$colName][$currency][$priceList] = $colIdx;
                            break;
                        case 'Name':
                        case 'CategoryPath':
                        case 'PrimeCost':
                        case 'Description':
                        case 'ShortDescription':
                        case 'Publish':
                        case 'AttributeSet':
                            $parameter = strtoupper(trim($matches[1][0]));
                            if (!$parameter) {
                                $errorTxt .= "No required parameter for column '{$headerCell}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                            }
                            if (array_key_exists($parameter, $this->colMap[$colName])) {
                                $errorTxt .= "Duplicate column '{$headerCell}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                            }

                            $this->colMap[$colName][$parameter] = $colIdx;
                            break;
                        default:
                            if (null === $this->colMap[$colName]) {
                                $this->colMap[$colName] = $colIdx;
                            } else {
                                $errorTxt .= "Duplicate column '{$headerCell}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                            }
                    }
                }
            }
        }

        foreach ($this->requiredColumns as $requiredColumn) {
            if (empty($this->colMap[$requiredColumn]) && 0 !== $this->colMap[$requiredColumn]) {
                $errorTxt .= "Required column '{$requiredColumn}' not found!" . PHP_EOL;
            }
        }

        if ($errorTxt) {
            throw new Error($errorTxt);
        }
    }

    private function insertElemWithLocale(DOMElement $productEl, $rowIdx, $fieldName, $isBoolean)
    {
        foreach ($this->colMap[$fieldName] as $locale => $colIdx) {
            /** @noinspection PhpIllegalArrayKeyTypeInspection */
            $val = array_key_exists($colIdx, $this->rows[$rowIdx]) ? trim($this->rows[$rowIdx][$colIdx]) : '';
            if ($isBoolean) {
                $val = filter_var($val, FILTER_VALIDATE_BOOLEAN) ? '1' : '0';
            }
            if ('' != $val) {
                $newElem = $this->domDocument->createElement($fieldName, $isBoolean ? $val : null);
                $newElem->setAttribute('locale', $locale);

                if (!$isBoolean) {
                    $newElem->appendChild($this->domDocument->createCDATASection($val));
                }
                $productEl->appendChild($newElem);
            }
        }
    }

    private function insertElemPrice(DOMElement $productEl, $rowIdx, $fieldName)
    {
        foreach ($this->colMap[$fieldName] as $currency => $colIdx) {
            if (is_array($colIdx)) {
                foreach ($colIdx as $priceListName => $idx) {
                    $val = array_key_exists($idx, $this->rows[$rowIdx]) ? trim($this->rows[$rowIdx][$idx]) : '';
                    if ('' !== $val) {
                        $newElem = $this->domDocument->createElement($fieldName, $val);
                        $newElem->setAttribute('currency', $currency);
                        if ($priceListName) {
                            $newElem->setAttribute('priceListName', $priceListName);
                        }
                        $productEl->appendChild($newElem);
                    }
                }
            } else {
                /** @noinspection PhpIllegalArrayKeyTypeInspection */
                $val = trim($this->rows[$rowIdx][$colIdx]);
                if ('' !== $val) {
                    $newElem = $this->domDocument->createElement($fieldName, $val);
                    $newElem->setAttribute('currency', $currency);
                    $productEl->appendChild($newElem);
                }
            }

        }
    }

    /**
     * @param $rowIdx
     * @return DOMElement|null
     */
    private function getProductElem($rowIdx)
    {
        $errorTxt = '';

        $colIdx = $this->colMap['export'];
        $row = &$this->rows[$rowIdx];
        $val = trim($row[$colIdx]);

        if (filter_var($val, FILTER_VALIDATE_BOOLEAN)) {
            foreach ($this->requiredColumns as $requiredColumn) {
                $colIdx = $this->colMap[$requiredColumn];
                if (is_array($colIdx)) {
                    $valObj = (object)['val' => '', 'row' => &$row];
                    array_walk_recursive($colIdx, function ($item, $key, $userdata) {
                        $valInside = trim($userdata->row[$item]);
                        if ($valInside) {
                            $userdata->val = $valInside;
                        }
                    }, $valObj);
                    $val = $valObj->val;
                } else {
                    $val = trim($row[$colIdx]);
                }

                if ('' === $val) {
                    $errorTxt .= "    Required column '{$requiredColumn}' has no value" . PHP_EOL;
                }
            }

            $code = null;
            $codeUpper = null;
            $barcode = null;
            $barcodeUpper = null;
            $images = [];

            foreach ($this->colMap as $colName => $colIdx) {
                if (null === $colIdx) {
                    continue;
                }
                switch ($colName) {
                    case 'code':
                        /** @noinspection PhpIllegalArrayKeyTypeInspection */
                        $code = trim($row[$colIdx]);
                        $codeUpper = $this->codeCaseSensitive ? $code : mb_strtoupper($code);
                        if (in_array($codeUpper, $this->processedCodes)) {
                            /** @noinspection PhpIllegalArrayKeyTypeInspection */
                            $errorTxt .= "    Column {$this->getColNameFromNumber($colIdx)} - duplicate code '{$row[$colIdx]}'" . PHP_EOL;
                        }
                        break;
                    case 'Barcode':
                        /** @noinspection PhpIllegalArrayKeyTypeInspection */
                        $barcode = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
                        if ('' !== $barcode) {
                            $barcodeUpper = mb_strtoupper($barcode);
                            if (in_array($barcodeUpper, $this->processedBarcodes)) {
                                /** @noinspection PhpIllegalArrayKeyTypeInspection */
                                $errorTxt .= "    Column {$this->getColNameFromNumber($colIdx)} - duplicate barcode '{$row[$colIdx]}'" . PHP_EOL;
                            }
                        }
                        break;
                    case 'Image':
                        foreach ($colIdx as $idx) {
                            $val = array_key_exists($idx, $row) ? trim($row[$idx]) : '';
                            if ('' !== $val) {
                                if ($val = filter_var($val, FILTER_VALIDATE_URL)) {
                                    $images[] = $val;
                                } else {
                                    $errorTxt .= "    Column {$this->getColNameFromNumber($idx)} - invalid url '{$row[$idx]}'" . PHP_EOL;
                                }
                            }
                        }
                        break;
                    case 'Weight':
                        /** @noinspection PhpIllegalArrayKeyTypeInspection */
                        $val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
                        if ('' !== $val) {
                            if (false === filter_var($val, FILTER_VALIDATE_INT)) {
                                /** @noinspection PhpIllegalArrayKeyTypeInspection */
                                $errorTxt .= "    Column {$this->getColNameFromNumber($colIdx)} - invalid integer '{$row[$colIdx]}'" . PHP_EOL;
                            }
                        }
                        break;
                    case 'Tax':
                    case 'PrimeCost':
                    case 'Price':
                    case 'OldPrice':
                    case 'Quantity':
                        if (is_array($colIdx)) {
                            $valObj = (object)['val' => '', 'row' => &$row, 'errorTxt' => &$errorTxt];
                            array_walk_recursive($colIdx, function ($item, $key, $userdata) {
                                $valInside = array_key_exists($item, $userdata->row) ? trim($userdata->row[$item]) : '';
                                if ('' !== $valInside) {
                                    if (false === filter_var($valInside, FILTER_VALIDATE_FLOAT)) {
                                        $userdata->errorTxt .= "    Column {$this->getColNameFromNumber($item)} - invalid float '{$userdata->row[$item]}'" . PHP_EOL;
                                    }
                                }
                            }, $valObj);
                        } else {
                            /** @noinspection PhpIllegalArrayKeyTypeInspection */
                            $val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
                            if ('' !== $val) {
                                if (false === filter_var($val, FILTER_VALIDATE_FLOAT)) {
                                    /** @noinspection PhpIllegalArrayKeyTypeInspection */
                                    $errorTxt .= "    Column {$this->getColNameFromNumber($colIdx)} - invalid float '{$row[$colIdx]}'" . PHP_EOL;
                                }
                            }
                        }
                        break;
                }

            }


            if ($errorTxt) {
                echo 'ERROR row number ', $rowIdx + 1, PHP_EOL, $errorTxt;
                $this->invalidRows++;
                return null;
            } else {
                $productEl = $this->domDocument->createElement('Product');

                $productEl->setAttribute('code', $code);

                if ($barcode) {
                    $newElem = $this->domDocument->createElement('Barcode');
                    $newElem->appendChild($this->domDocument->createCDATASection($barcode));
                    $productEl->appendChild($newElem);
                }

                $this->insertElemWithLocale($productEl, $rowIdx, 'CategoryPath', false);
                $this->insertElemWithLocale($productEl, $rowIdx, 'Name', false);
                $this->insertElemWithLocale($productEl, $rowIdx, 'Description', false);
                $this->insertElemWithLocale($productEl, $rowIdx, 'ShortDescription', false);

                $val = trim($row[$this->colMap['Tax']]);
                $newElem = $this->domDocument->createElement('Tax', $val);
                $productEl->appendChild($newElem);


                $this->insertElemPrice($productEl, $rowIdx, 'PrimeCost');
                $this->insertElemPrice($productEl, $rowIdx, 'Price');
                $this->insertElemPrice($productEl, $rowIdx, 'OldPrice');

                $val = trim($row[$this->colMap['Quantity']]);
                $newElem = $this->domDocument->createElement('Quantity', $val);
                $productEl->appendChild($newElem);

                $colIdx = $this->colMap['Weight'];
                if ($val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '') {
                    $newElem = $this->domDocument->createElement('Weight', $val);
                    $productEl->appendChild($newElem);
                }

                $this->insertElemWithLocale($productEl, $rowIdx, 'AttributeSet', false);

                foreach ($this->attrCols as $attrColIdx => $attrNames) {
                    $attrNameElems = [];
                    $attrValueElems = [];
                    $attrLocaleMap = [];

                    $attrVals = array_key_exists($attrColIdx, $row) ? explode($this->attrLangSeparator, $row[$attrColIdx]) : [];

                    foreach ($attrNames as $attrLocale => $attrName) {
                        $attrNameElem = $this->domDocument->createElement('Name');
                        $attrNameElem->setAttribute('locale', $attrLocale);
                        $attrNameElem->appendChild($this->domDocument->createCDATASection($attrName));
                        $attrNameElems[] = $attrNameElem;
                        $attrLocaleMap[] = $attrLocale;
                    }

                    $attrLocalesLength = count($attrLocaleMap);
                    for ($i = 0; $i < $attrLocalesLength; $i++) {
                        $attrValue = array_key_exists($i, $attrVals) ? trim($attrVals[$i]) : '';

                        $attrValueElem = $this->domDocument->createElement('Value');
                        $attrValueElem->setAttribute('locale', $attrLocaleMap[$i]);
                        if ('' !== $attrValue) {
                            $attrValueElem->appendChild($this->domDocument->createCDATASection($attrValue));
                            $attrValueElems[$i] = $attrValueElem;
                        }
                    }

                    $arr_keys = array_keys($attrValueElems);
                    if (!empty($arr_keys)) {
                        $newElem = $this->domDocument->createElement('Attribute');
                        foreach ($arr_keys as $arr_key) {
                            $newElem->appendChild($attrNameElems[$arr_key]);
                        }
                        foreach ($arr_keys as $arr_key) {
                            $newElem->appendChild($attrValueElems[$arr_key]);
                        }
                        $productEl->appendChild($newElem);
                    }
                }


                foreach ($images as $image) {
                    $newElem = $this->domDocument->createElement('Image');
                    $newElem->appendChild($this->domDocument->createCDATASection($image));
                    $productEl->appendChild($newElem);
                }

                $this->insertElemWithLocale($productEl, $rowIdx, 'Publish', true);


                $this->exportedRows++;
                $this->processedCodes[] = $codeUpper;
                if ($barcodeUpper) {
                    $this->processedBarcodes[] = $barcodeUpper;
                }
                return $productEl;
            }
        } else {
            $this->skippedRows++;
            return null;
        }
    }

    /**
     * @param $spreadsheetId
     * @param $sheetName
     * @param $outputFilename
     * @throws Google_Exception
     */
    public function generateProducts($spreadsheetId, $sheetName, $outputFilename)
    {
        $client = $this->getClient();
        $service = new Google_Service_Sheets($client);

        echo 'Spreadsheet ID: ', $spreadsheetId, PHP_EOL, 'Sheet: ', $sheetName, PHP_EOL;

        $result = $service->spreadsheets_values->get($spreadsheetId, $sheetName);
        $this->rows = $result->getValues();

        $this->rowsCount = count($this->rows);
        echo 'Total rows: ', $this->rowsCount - 1, PHP_EOL;

        if ($this->rowsCount > 0) {
            $this->mapColumns();

            $this->codeCaseSensitive = filter_var(getenv('CODE_CASE_SENSITIVE'), FILTER_VALIDATE_BOOLEAN);

            $this->processedCodes = [];
            $this->processedBarcodes = [];
            $this->exportedRows = 0;
            $this->skippedRows = 0;
            $this->invalidRows = 0;

            $this->domDocument = new DOMDocument();
//            $this->domDocument->formatOutput = true;
            $rootEl = $this->domDocument->createElement('Products');
            $rootEl->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
            $rootEl->setAttribute('xsi:noNamespaceSchemaLocation', 'http://xsd.verskis.lt/product-import/products.xsd');
            $this->domDocument->appendChild($rootEl);

            for ($i = 1; $i < $this->rowsCount; $i++) {
                if ($productElem = $this->getProductElem($i)) {
                    $rootEl->appendChild($productElem);
                }
            }

            if ($this->exportedRows > 0) {
                $this->domDocument->save($outputFilename);
            }

            echo PHP_EOL, 'COMPLETED', PHP_EOL, 'Exported: ', $this->exportedRows, PHP_EOL, 'Skipped: ', $this->skippedRows, PHP_EOL, 'Invalid: ', $this->invalidRows, PHP_EOL;


        } else {
            throw new Error("No header row!");
        }
    }
}

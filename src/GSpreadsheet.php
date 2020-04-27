<?php

namespace SVDVerskisLT;

use DOMDocument;
use DOMElement;
use Exception;
use Google_Client;
use Google_Exception;
use Google_Service_Sheets;
use Google_Service_Sheets_ValueRange;
use RuntimeException;

class GSpreadsheet
{
    private $requiredColumns = ['export', 'code', 'CategoryPath', 'Name', 'Tax', 'PrimeCost', 'Price', 'Quantity'];
    private $optsColMap;
    private $prodColMap;
    private $attrCols;
    private $optAttrCols;
    private $attrLangSeparator;
    private $codeCaseSensitive;
    private $prodRows;
    private $opts;
    /** @var DOMDocument */
    private $domDocument;
    private $processedCodes;
    private $processedBarcodes;
    private $exportedRows;
    private $skippedRows;
    private $invalidRows;
    private $exportedOptions;
    private $skippedOptions;
    private $invalidOptions;

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

    private function mapOptsColumns(&$optsRows)
    {
        $errorTxt = '';
        $totalOptsCols = count($optsRows[0]);
        $this->optsColMap = [
            'productCode' => null,
            'optionCode' => null,
            'exportOption' => null,
            'Barcode' => null,
            'PrimeCost' => [],
            'Price' => [],
            'OldPrice' => [],
            'Quantity' => null,
            'Weight' => null,
            'Publish' => []
        ];

        for ($colIdx = 0; $colIdx < $totalOptsCols; $colIdx++) {
            $headerCell = $optsRows[0][$colIdx];

            $matches = [];

            $colName = $headerCell;
            if (preg_match_all('/\[([^\[]+)]/', $headerCell, $matches)) {
                $colName = substr($headerCell, 0, strpos($headerCell, '['));
            }

            if (!array_key_exists($colName, $this->optsColMap)) {
                if (in_array($colName, $this->optAttrCols)) {
                    $locale = strtoupper(trim($matches[1][0]));
                    if (!$locale) {
                        $errorTxt .= "Locale not specified for column '{$headerCell}' in options sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                    }
                    if (array_key_exists($colName, $this->optsColMap) && array_key_exists($locale, $this->optsColMap[$colName])) {
                        $errorTxt .= "Duplicate column '{$headerCell}' in options sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                    }

                    $this->optsColMap[$colName][$locale] = $colIdx;
                } else {
                    echo "WARNING: Unknown option column name '{$headerCell}' in options sheet column {$this->getColNameFromNumber($colIdx)}", PHP_EOL;
                }
            } else {
                switch ($colName) {
                    case 'Price':
                    case 'OldPrice':
                        $currency = strtoupper(trim($matches[1][0]));
                        $priceList = array_key_exists(1, $matches[1]) ? strtoupper(trim($matches[1][1])) : '';
                        if (array_key_exists($currency, $this->optsColMap[$colName]) && array_key_exists($priceList, $this->optsColMap[$colName][$currency])) {
                            $errorTxt .= "Duplicate column '{$headerCell}' in options sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                        }
                        $this->optsColMap[$colName][$currency][$priceList] = $colIdx;
                        break;
                    case 'productCode':
                    case 'optionCode':
                    case 'exportOption':
                    case 'Barcode':
                    case 'Quantity':
                    case 'Weight':
                        if (null === $this->optsColMap[$colName]) {
                            $this->optsColMap[$colName] = $colIdx;
                        } else {
                            $errorTxt .= "Duplicate column '{$headerCell}' in options sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                        }
                        break;
                    default:
                        $parameter = strtoupper(trim($matches[1][0]));
                        if (!$parameter) {
                            $errorTxt .= "No required parameter for column '{$headerCell}' in options sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                        }
                        if (array_key_exists($parameter, $this->optsColMap[$colName])) {
                            $errorTxt .= "Duplicate column '{$headerCell}' in options sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                        }

                        $this->optsColMap[$colName][$parameter] = $colIdx;
                        break;
                }
            }

        }


    }

    private function mapProdColumns()
    {
        $errorTxt = '';

        $this->attrLangSeparator = getenv('ATTR_LANG_SEPARATOR') ?: '¦';

        $totalCols = count($this->prodRows[0]);

        $this->attrCols = [];
        $this->optAttrCols = [];
        $attrNamesUniq = [];
        $optAttrNamesUniq = [];

        $this->prodColMap = [
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
            $headerCell = $this->prodRows[0][$colIdx];

            if (!strncasecmp('attr:', $headerCell, 5)) {
                $rawAttrNames = explode($this->attrLangSeparator, substr($headerCell, 5));

                foreach ($rawAttrNames as $rawAttrName) {
                    $matches = [];
                    if (preg_match('/^([^\[]+)\[([^\[]+)]$/', trim($rawAttrName), $matches)) {
                        $attrName = trim($matches[1]);
                        $attrNameUpper = mb_strtoupper($attrName);

                        $parameter = strtoupper($matches[2]);
                        if (!array_key_exists($parameter, $attrNamesUniq)) $attrNamesUniq[$parameter] = [];

                        if (in_array($attrNameUpper, $attrNamesUniq[$parameter]))
                            $errorTxt .= "Duplicate attribute name '{$rawAttrName}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;

                        $attrNamesUniq[$parameter][] = $attrNameUpper;

                        if (!array_key_exists($colIdx, $this->attrCols)) $this->attrCols[$colIdx] = [];
                        if (!array_key_exists($parameter, $this->attrCols[$colIdx])) $this->attrCols[$colIdx][$parameter] = $attrName;

                    } else {
                        $errorTxt .= "Invalid attribute name '{$rawAttrName}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                    }
                }
            } elseif (!strncasecmp('optAttr:', $headerCell, 8)) {
                $optAttrCol = trim(substr($headerCell, 8));

                if ($optAttrCol) {
                    $optAttrColUpper = mb_strtoupper($optAttrCol);

                    if (in_array($optAttrColUpper, $optAttrNamesUniq))
                        $errorTxt .= "Duplicate option attribute name '{$optAttrCol}' on products sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;

                    $optAttrNamesUniq[] = $optAttrColUpper;

                    $this->optAttrCols[$colIdx] = $optAttrCol;
                } else {
                    $errorTxt .= "Invalid option attribute name '{$optAttrCol}' on products sheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                }
            } else {
                $matches = [];

                $colName = $headerCell;
                if (preg_match_all('/\[([^\[]+)]/', $headerCell, $matches)) {
                    $colName = substr($headerCell, 0, strpos($headerCell, '['));
                }

                if (!array_key_exists($colName, $this->prodColMap)) {
                    echo "WARNING: Unknown column name '{$headerCell}' in column {$this->getColNameFromNumber($colIdx)}", PHP_EOL;
                } else {
                    switch ($colName) {
                        case 'Image':
                            $this->prodColMap[$colName][] = $colIdx;
                            break;
                        case 'Price':
                        case 'OldPrice':
                            $currency = strtoupper(trim($matches[1][0]));
                            $priceList = array_key_exists(1, $matches[1]) ? strtoupper(trim($matches[1][1])) : '';
                            if (array_key_exists($currency, $this->prodColMap[$colName]) && array_key_exists($priceList, $this->prodColMap[$colName][$currency])) {
                                $errorTxt .= "Duplicate column '{$headerCell}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                            }
                            $this->prodColMap[$colName][$currency][$priceList] = $colIdx;
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
                            if (array_key_exists($parameter, $this->prodColMap[$colName])) {
                                $errorTxt .= "Duplicate column '{$headerCell}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                            }

                            $this->prodColMap[$colName][$parameter] = $colIdx;
                            break;
                        default:
                            if (null === $this->prodColMap[$colName]) {
                                $this->prodColMap[$colName] = $colIdx;
                            } else {
                                $errorTxt .= "Duplicate column '{$headerCell}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                            }
                    }
                }
            }
        }

        foreach ($this->requiredColumns as $requiredColumn) {
            if (empty($this->prodColMap[$requiredColumn]) && 0 !== $this->prodColMap[$requiredColumn]) {
                $errorTxt .= "Required column '{$requiredColumn}' not found!" . PHP_EOL;
            }
        }

        if ($errorTxt) {
            throw new RuntimeException($errorTxt);
        }
    }

    private function insertOptionAttributeValues(DOMElement $attributeEl, $fieldName, &$row)
    {
        $firstLocale = null;
        $atLeastOneValue = false;
        foreach ($this->optsColMap[$fieldName] as $locale => $colIdx) {
            if (null === $firstLocale) $firstLocale = $locale;
            $val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
            if ('' !== $val) {
                $valueElem = $this->domDocument->createElement('Value');
                $valueElem->setAttribute('locale', $locale);
                $valueElem->appendChild($this->domDocument->createCDATASection($val));
                $attributeEl->appendChild($valueElem);
                $atLeastOneValue = true;
            }
        }

        if (!$atLeastOneValue) {
            $valueElem = $this->domDocument->createElement('Value');
            $valueElem->setAttribute('locale', $firstLocale);
            $valueElem->appendChild($this->domDocument->createCDATASection(''));
            $attributeEl->appendChild($valueElem);
        }
    }

    private function insertElemWithLocale(DOMElement $productEl, $rowIdx, $fieldName, $isBoolean)
    {
        foreach ($this->prodColMap[$fieldName] as $locale => $colIdx) {
            $val = array_key_exists($colIdx, $this->prodRows[$rowIdx]) ? trim($this->prodRows[$rowIdx][$colIdx]) : '';
            if ($isBoolean) {
                $val = filter_var($val, FILTER_VALIDATE_BOOLEAN) ? '1' : '0';
            }
            if ('' !== $val) {
                $newElem = $this->domDocument->createElement($fieldName, $isBoolean ? $val : null);
                $newElem->setAttribute('locale', $locale);

                if (!$isBoolean) {
                    $newElem->appendChild($this->domDocument->createCDATASection($val));
                }
                $productEl->appendChild($newElem);
            }
        }
    }

    private function insertOptElemWithLocale(DOMElement $optionEl, $row, $fieldName, $isBoolean)
    {
        foreach ($this->optsColMap[$fieldName] as $locale => $colIdx) {
            $val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
            if ('' !== $val) {
                if ($isBoolean) {
                    $val = filter_var($val, FILTER_VALIDATE_BOOLEAN) ? '1' : '0';
                }
                $newElem = $this->domDocument->createElement($fieldName, $isBoolean ? $val : null);
                $newElem->setAttribute('locale', $locale);

                if (!$isBoolean) {
                    $newElem->appendChild($this->domDocument->createCDATASection($val));
                }
                $optionEl->appendChild($newElem);
            }
        }
    }


    private function insertOptElemPrice(DOMElement $optionEl, $row, $fieldName)
    {
        foreach ($this->optsColMap[$fieldName] as $currency => $colIdx) {
            if (is_array($colIdx)) {
                foreach ($colIdx as $priceListName => $idx) {
                    $val = array_key_exists($idx, $row) ? trim($row[$idx]) : '';
                    if ('' !== $val) {
                        $newElem = $this->domDocument->createElement($fieldName, $val);
                        $newElem->setAttribute('currency', $currency);
                        if ($priceListName) {
                            $newElem->setAttribute('priceListName', $priceListName);
                        }
                        $optionEl->appendChild($newElem);
                    }
                }
            } else {
                $val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
                if ('' !== $val) {
                    $newElem = $this->domDocument->createElement($fieldName, $val);
                    $newElem->setAttribute('currency', $currency);
                    $optionEl->appendChild($newElem);
                }
            }
        }
    }


    private function insertElemPrice(DOMElement $productEl, $rowIdx, $fieldName)
    {
        foreach ($this->prodColMap[$fieldName] as $currency => $colIdx) {
            if (is_array($colIdx)) {
                foreach ($colIdx as $priceListName => $idx) {
                    $val = array_key_exists($idx, $this->prodRows[$rowIdx]) ? trim($this->prodRows[$rowIdx][$idx]) : '';
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
                $val = trim($this->prodRows[$rowIdx][$colIdx]);
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
     * @noinspection PhpUnusedParameterInspection
     */
    private function getProductElem($rowIdx)
    {
        $errorTxt = '';

        $colIdx = $this->prodColMap['export'];
        $row = &$this->prodRows[$rowIdx];
        $val = trim($row[$colIdx]);

        if (filter_var($val, FILTER_VALIDATE_BOOLEAN)) {
            foreach ($this->requiredColumns as $requiredColumn) {
                $colIdx = $this->prodColMap[$requiredColumn];
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

            foreach ($this->prodColMap as $colName => $colIdx) {
                if (null === $colIdx) {
                    continue;
                }
                switch ($colName) {
                    case 'code':
                        $code = trim($row[$colIdx]);
                        $codeUpper = $this->codeCaseSensitive ? $code : mb_strtoupper($code);
                        if (in_array($codeUpper, $this->processedCodes)) {
                            $errorTxt .= "    Column {$this->getColNameFromNumber($colIdx)} - duplicate code '{$row[$colIdx]}'" . PHP_EOL;
                        }
                        break;
                    case 'Barcode':
                        $barcode = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
                        if ('' !== $barcode) {
                            $barcodeUpper = mb_strtoupper($barcode);
                            if (in_array($barcodeUpper, $this->processedBarcodes)) {
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
                        $val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
                        if ('' !== $val) {
                            if (false === filter_var($val, FILTER_VALIDATE_INT)) {
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
                            $val = array_key_exists($colIdx, $row) ? trim($row[$colIdx]) : '';
                            if ('' !== $val) {
                                if (false === filter_var($val, FILTER_VALIDATE_FLOAT)) {
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

                $val = trim($row[$this->prodColMap['Tax']]);
                $newElem = $this->domDocument->createElement('Tax', $val);
                $productEl->appendChild($newElem);


                $this->insertElemPrice($productEl, $rowIdx, 'PrimeCost');
                $this->insertElemPrice($productEl, $rowIdx, 'Price');
                $this->insertElemPrice($productEl, $rowIdx, 'OldPrice');

                $val = trim($row[$this->prodColMap['Quantity']]);
                $newElem = $this->domDocument->createElement('Quantity', $val);
                $productEl->appendChild($newElem);

                $colIdx = $this->prodColMap['Weight'];
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

                    foreach ($attrNames as $attrLocale => $optAttrName) {
                        $attrNameElem = $this->domDocument->createElement('Name');
                        $attrNameElem->setAttribute('locale', $attrLocale);
                        $attrNameElem->appendChild($this->domDocument->createCDATASection($optAttrName));
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

                $this->processedCodes[] = $codeUpper;
                if ($barcodeUpper) $this->processedBarcodes[] = $barcodeUpper;

                $optAttrColNames = [];

                foreach ($this->optAttrCols as $optAttrColIdx => $optAttrCol) {
                    $optAttrColValue = trim($row[$optAttrColIdx]);
                    $rawOptAttrNames = $optAttrColValue ? explode($this->attrLangSeparator, $optAttrColValue) : [];

                    if ($rawOptAttrNames) {
                        $optionsAttributeElem = $this->domDocument->createElement('OptionsAttribute');
                        $optAttrColNames[] = $optAttrCol;
                        foreach ($rawOptAttrNames as $rawOptAttrName) {
                            $matches = [];
                            if (preg_match('/^([^\[]+)\[([^\[]+)]$/', trim($rawOptAttrName), $matches)) {
                                $optAttrName = trim($matches[1]);
                                $optAttrLocale = strtoupper($matches[2]);

                                $newElem = $this->domDocument->createElement('Name');
                                $newElem->setAttribute('locale', $optAttrLocale);
                                $newElem->appendChild($this->domDocument->createCDATASection($optAttrName));
                                $optionsAttributeElem->appendChild($newElem);
                            } else {
                                $this->invalidRows++;
                                echo 'ERROR row number ', $rowIdx + 1, " Invalid options attribute name '{$rawOptAttrName}' on spreadsheet column {$this->getColNameFromNumber($colIdx)}" . PHP_EOL;
                                return null;
                            }
                        }
                        $productEl->appendChild($optionsAttributeElem);
                    }
                }

                if ($optAttrColNames) {
                    if (!$this->addOptions($productEl, $codeUpper, $optAttrColNames)) {
                        $this->invalidRows++;
                        echo 'ERROR row number ', $rowIdx + 1, " Product has no options specified" . PHP_EOL;
                        return null;
                    }
                }

                foreach ($images as $image) {
                    $newElem = $this->domDocument->createElement('Image');
                    $newElem->appendChild($this->domDocument->createCDATASection($image));
                    $productEl->appendChild($newElem);
                }

                $this->insertElemWithLocale($productEl, $rowIdx, 'Publish', true);


                $this->exportedRows++;

                return $productEl;
            }
        } else {
            $this->skippedRows++;
            return null;
        }
    }

    /**
     * @param DOMElement $productEl
     * @param string $productCodeUpper
     * @param array $optAttrColNames
     * @return int
     */
    private function addOptions($productEl, $productCodeUpper, $optAttrColNames)
    {
        $addedOptions = 0;

        if (array_key_exists($productCodeUpper, $this->opts)) {
            foreach ($this->opts[$productCodeUpper] as $optionCodeUpper => $optionRow) {
                $colIdx = $this->optsColMap['optionCode'];
                $optionCode = trim($optionRow[$colIdx]);

                if (in_array($optionCodeUpper, $this->processedCodes)) {
                    echo "ERROR Duplicate option code '{$optionCode}'" . PHP_EOL;
                    $this->invalidOptions++;
                    continue;
                }

                $optionElem = $this->domDocument->createElement('Option');
                $optionElem->setAttribute('code', $optionCode);
                foreach ($optAttrColNames as $optAttrColName) {
                    $attributeElem = $this->domDocument->createElement('Attribute');
                    $this->insertOptionAttributeValues($attributeElem, $optAttrColName, $optionRow);
                    $optionElem->appendChild($attributeElem);
                }

                $this->insertOptElemPrice($optionElem, $optionRow, 'PrimeCost');
                $this->insertOptElemPrice($optionElem, $optionRow, 'Price');
                $this->insertOptElemPrice($optionElem, $optionRow, 'OldPrice');

                if (array_key_exists('Quantity', $this->optsColMap)) {
                    $colIdx = $this->optsColMap['Quantity'];
                    if ($val = array_key_exists($colIdx, $optionRow) ? trim($optionRow[$colIdx]) : '') {
                        $newElem = $this->domDocument->createElement('Quantity', $val);
                        $optionElem->appendChild($newElem);
                    }
                }

                if (array_key_exists('Weight', $this->optsColMap)) {
                    $colIdx = $this->optsColMap['Weight'];
                    if ($val = array_key_exists($colIdx, $optionRow) ? trim($optionRow[$colIdx]) : '') {
                        $newElem = $this->domDocument->createElement('Weight', $val);
                        $optionElem->appendChild($newElem);
                    }
                }

                if (array_key_exists('Barcode', $this->optsColMap)) {
                    $colIdx = $this->optsColMap['Barcode'];
                    if ($val = array_key_exists($colIdx, $optionRow) ? trim($optionRow[$colIdx]) : '') {
                        $barcodeUpper = mb_strtoupper($val);
                        if (in_array($barcodeUpper, $this->processedBarcodes)) {
                            echo "ERROR Duplicate option barcode '{$val}'" . PHP_EOL;
                            $this->invalidOptions++;
                            continue;
                        }

                        $newElem = $this->domDocument->createElement('Barcode');
                        $newElem->appendChild($this->domDocument->createCDATASection($val));
                        $optionElem->appendChild($newElem);
                        $this->processedBarcodes[] = $barcodeUpper;
                    }
                }

                if (array_key_exists('Publish', $this->optsColMap)) {
                    $this->insertOptElemWithLocale($optionElem, $optionRow, 'Publish', true);
                }

                $this->processedCodes[] = $optionCodeUpper;

                $productEl->appendChild($optionElem);
                $this->exportedOptions++;
                $addedOptions++;

            }
        }

        return $addedOptions;
    }

    /**
     * @param Google_Service_Sheets_ValueRange $optionsResult
     */
    private function prepareOptionsSheet($optionsResult)
    {
        $optionsRowsCount = $optionsResult->count();
        echo 'Total options rows: ', $optionsRowsCount - 1, PHP_EOL;
        $this->opts = [];

        if ($optionsRowsCount > 0) {
            $rawOptsRows = $optionsResult->getValues();
            $this->mapOptsColumns($rawOptsRows);

            $productCodeIdx = array_key_exists('productCode', $this->optsColMap) ? $this->optsColMap['productCode'] : null;
            if (null === $productCodeIdx) {
                echo 'ERROR: Column "productCode" not found in options sheet', PHP_EOL;
                return;
            }
            $optionCodeIdx = array_key_exists('optionCode', $this->optsColMap) ? $this->optsColMap['optionCode'] : null;
            if (null === $optionCodeIdx) {
                echo 'ERROR: Column "optionCode" not found in options sheet', PHP_EOL;
                return;
            }

            for ($i = 1; $i < $optionsRowsCount; $i++) {
                $optionRow = $rawOptsRows[$i];
                $productCode = array_key_exists($productCodeIdx, $optionRow) ? trim($optionRow[$productCodeIdx]) : '';
                $optionCode = array_key_exists($optionCodeIdx, $optionRow) ? trim($optionRow[$optionCodeIdx]) : '';

                if ($productCode) {
                    if (!$this->codeCaseSensitive) $productCode = mb_strtoupper($productCode);
                } else {
                    echo 'ERROR options sheet row ' . ($i + 1) . ' "productCode" not specified', PHP_EOL;
                    $this->invalidOptions++;
                    continue;
                }

                if ($optionCode) {
                    if (!$this->codeCaseSensitive) $optionCode = mb_strtoupper($optionCode);
                } else {
                    echo 'ERROR options sheet row ' . ($i + 1) . ' "optionCode" not specified', PHP_EOL;
                    $this->invalidOptions++;
                    continue;
                }

                $colIdx = $this->optsColMap['exportOption'];
                $exportOption = array_key_exists($colIdx, $optionRow) ? trim($optionRow[$colIdx]) : false;

                if (filter_var($exportOption, FILTER_VALIDATE_BOOLEAN)) {
                    if (array_key_exists($productCode, $this->opts) && array_key_exists($optionCode, $this->opts[$productCode])) {
                        echo 'ERROR options sheet row ' . ($i + 1) . ' duplicate "optionCode" for product', PHP_EOL;
                    } else {
                        $this->opts[$productCode][$optionCode] = &$rawOptsRows[$i];
                    }
                } else {
                    $this->skippedOptions++;
                }
            }

        } else {
            throw new RuntimeException("No header row in options sheet!");
        }
    }

    /**
     * @param $spreadsheetId
     * @param $sheetName
     * @param $outputFilename
     * @param $optionsSheetName
     * @throws Google_Exception
     */
    public function generateProducts($spreadsheetId, $sheetName, $outputFilename, $optionsSheetName = null)
    {
        $client = $this->getClient();
        $service = new Google_Service_Sheets($client);

        echo 'Spreadsheet ID: ', $spreadsheetId, PHP_EOL, 'Sheet: ', $sheetName, PHP_EOL;
        if ($optionsSheetName) echo 'Options sheet: ', $optionsSheetName, PHP_EOL;

        $result = $service->spreadsheets_values->get($spreadsheetId, $sheetName);
        $rowsCount = $result->count();
        echo 'Total rows: ', $rowsCount - 1, PHP_EOL;

        $this->prodRows = $result->getValues();
        if ($rowsCount > 0) {
            $this->codeCaseSensitive = filter_var(getenv('CODE_CASE_SENSITIVE'), FILTER_VALIDATE_BOOLEAN);
            $this->processedCodes = [];
            $this->processedBarcodes = [];
            $this->exportedRows = 0;
            $this->skippedRows = 0;
            $this->invalidRows = 0;
            $this->exportedOptions = 0;
            $this->skippedOptions = 0;
            $this->invalidOptions = 0;

            $this->mapProdColumns();
            if ($optionsSheetName) {
                $result = $service->spreadsheets_values->get($spreadsheetId, $optionsSheetName);
                $this->prepareOptionsSheet($result);
            }

            $this->domDocument = new DOMDocument();
            $this->domDocument->formatOutput = filter_var(getenv('XML_FORMAT_OUTPUT'), FILTER_VALIDATE_BOOLEAN);
            $rootEl = $this->domDocument->createElement('Products');
            $rootEl->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
            $rootEl->setAttribute('xsi:noNamespaceSchemaLocation', 'http://xsd.verskis.lt/product-import/products.xsd');
            $this->domDocument->appendChild($rootEl);

            for ($i = 1; $i < $rowsCount; $i++) {
                if ($productElem = $this->getProductElem($i)) {
                    $rootEl->appendChild($productElem);
                }
            }

            if ($this->exportedRows > 0) {
                $this->domDocument->save($outputFilename);
            }

            echo PHP_EOL, 'COMPLETED', PHP_EOL,
            'Exported: ', $this->exportedRows, PHP_EOL,
            'Skipped: ', $this->skippedRows, PHP_EOL,
            'Invalid: ', $this->invalidRows, PHP_EOL,
            'Options exported: ', $this->exportedOptions, PHP_EOL,
            'Options skipped: ', $this->skippedOptions, PHP_EOL,
            'Options invalid: ', $this->invalidOptions, PHP_EOL;
        } else {
            throw new RuntimeException("No header row in product sheet!");
        }
    }
}

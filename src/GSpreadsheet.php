<?php

namespace SVDVerskisLT;

use Error;
use Exception;
use Google_Client;
use Google_Exception;
use Google_Service_Sheets;

class GSpreadsheet
{
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

    private function getColNameFromNumber($num) {
        $numeric = $num % 26;
        $letter = chr(65 + $numeric);
        $num2 = intval($num / 26);
        if ($num2 > 0) {
            return $this->getColNameFromNumber($num2 - 1) . $letter;
        } else {
            return $letter;
        }
    }

    /**
     * @param $outputfile
     * @throws Google_Exception
     */
    public function generateProducts($outputfile)
    {
        $client = $this->getClient();
        $service = new Google_Service_Sheets($client);

        $spreadsheetId = getenv('SPREADSHEET_ID');
        $sheetName = getenv('SHEETNAME');

        echo 'Spreadsheet ID: ', $spreadsheetId, PHP_EOL, 'Sheet: ', $sheetName, PHP_EOL;

        $result = $service->spreadsheets_values->get($spreadsheetId, $sheetName);
        $rows = $result->getValues();

        $totalRows = count($rows);
        echo 'Total rows: ', $totalRows - 1, PHP_EOL;

        if ($totalRows > 0) {
            $totalCols = count($rows[0]);

            $attrCols = [];
            $attrNamesUniq = [];

            for ($colIdx = 0; $colIdx < $totalCols; $colIdx++) {
                $headerCell = $rows[0][$colIdx];

                if (!strncasecmp('attr:', $headerCell, 5)) {
                    $rawAttrNames = explode('¦', substr($headerCell, 5));

                    foreach ($rawAttrNames as $rawAttrName) {
                        $matches = [];
                        if (preg_match('/^([^\[]+)\[([A-Za-z]{2,2})\]$/', trim($rawAttrName), $matches)) {
                            $attrName = trim($matches[1]);

                            $colNameUpper = mb_strtoupper($attrName);
                            $locale = strtoupper($matches[2]);
                            if( !array_key_exists($locale, $attrNamesUniq) ) {
                                $attrNamesUniq[$locale] = [];
                            }

                            if( in_array($colNameUpper, $attrNamesUniq[$locale]) ) {
                                throw new Error("Duplicate attribute name '{$rawAttrName}' on column {$this->getColNameFromNumber($colIdx)}");
                            }
                            $attrNamesUniq[$locale][] = $colNameUpper;

                            if( !array_key_exists($colIdx, $attrCols) )  {
                                $attrCols[$colIdx] = [];
                            }

                            if( !array_key_exists($locale, $attrCols[$colIdx]) ) {
                                $attrCols[$colIdx][$locale] = $attrName;
                            }
                        } else {
                            throw new Error("Invalid attribute name '{$rawAttrName}' on column {$this->getColNameFromNumber($colIdx)}");
                        }
                    }
                } else {





                }


//                echo $colIdx, '  -  ', $headerCell, PHP_EOL;

            }

            var_dump($attrCols);
        } else {
            throw new Error("No header row!");
        }


        $colMap = [
            'import' => null,
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
            'Weight' => [],
            'AttributeSet' => [],
            'Image' => [],
            'Publish' => []
        ];


//        var_dump($result->getValues());

//        echo $outputfile, PHP_EOL;
    }
}

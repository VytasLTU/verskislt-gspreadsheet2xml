<?php

namespace SVDVerskisLT;

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

        $client->setApplicationName('Verskis.LT XML generator');
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
                printf("Open the following link in your browser:\n%s\n", $authUrl);
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

    /**
     * @param $outputfile
     * @throws Google_Exception
     */
    public function generateProducts($outputfile)
    {
        $client = $this->getClient();
        $service = new Google_Service_Sheets($client);

        $spreadsheetId = '11fGBIq2eGJy5tRtq1OOBSrgrA6crLMAiIriWxzVaXIg';

        $result = $service->spreadsheets_values->get($spreadsheetId,'Sheet1');


        var_dump($result->getValues());

        echo $outputfile, PHP_EOL;
    }
}

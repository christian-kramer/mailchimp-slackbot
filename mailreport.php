<?php

error_reporting(E_ALL); ini_set('display_errors', 1);

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class MailChimp
{
    private $apikey;
    private $datacenter;
    public $basedir;
    private $context;

    function __construct()
    {
        $this->apikey = "";
        $this->datacenter = array_pop(explode('-', $this->apikey));

        $this->basedir = "../../../data/evanslarson.com/app/mailchimpprototype";

        $opts = array(
            'http'=>array(
            'method'=>"GET",
            'header'=>"Authorization: apikey $this->apikey"
            )
        );

        $this->context = stream_context_create($opts);
    }

    function query($string)
    {
        return json_decode(file_get_contents("http://$this->datacenter.api.mailchimp.com/3.0/$string", FALSE, $this->context));
    }
}

class Watson
{
    private $baseurl;
    private $username;
    private $password;
    private $clevercurl;
    public $querystring;

    function __construct()
    {
        $this->baseurl = "https://gateway.watsonplatform.net/natural-language-understanding/api/v1/analyze?version=2017-02-27";
        $this->username = "";
        $this->password = "";
    }
/*
    function clevercurl_init()
    {
        $this->clevercurl = curl_init();
        curl_setopt($this->clevercurl, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($this->clevercurl, CURLOPT_HTTPAUTH, CURLAUTH_ANY);
        curl_setopt($this->clevercurl, CURLOPT_USERPWD, "$this->username:$this->password");
        curl_setopt($this->clevercurl, CURLOPT_HTTPHEADER, array('Content-Type: application/json'));
    }
*/
    function analyze($firstname, $lastname, $jobtitle, $company)
    {
        if (!$jobtitle && !$company)
        {
            //var_dump("- no jobtitle or company for $firstname $lastname -");

            return array("title" => $jobtitle, "organization" => $company);
        }

        if (!$firstname)
        {
            $firstname = "null";
        }
        if (!$lastname)
        {
            $lastname = "null";
        }
        if (!$jobtitle)
        {
            $jobtitle = "null";
        }
        if (!$company)
        {
            $company = "null";
        }

        $this->querystring = "Here are 3 parameters: $firstname $lastname, $jobtitle, $company";

        //var_dump($this->querystring);

        //$this->clevercurl_init();

        $curl = curl_init();
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
        curl_setopt($curl, CURLOPT_HTTPAUTH, CURLAUTH_ANY);
        curl_setopt($curl, CURLOPT_USERPWD, "$this->username:$this->password");
        curl_setopt($curl, CURLOPT_HTTPHEADER, array('Content-Type: application/json'));
        curl_setopt($curl, CURLOPT_URL, $this->baseurl);
        curl_setopt($curl, CURLOPT_POST, 1);
        curl_setopt($curl, CURLOPT_POSTFIELDS, json_encode(array(
            "text" => $this->querystring,
            "features" => array(
                "entities" => array(
                    "sentiment" => true,
                    "limit" => 2
                )
            )
        )));

        $result = curl_exec($curl);
        $status_code = curl_getinfo($curl, CURLINFO_HTTP_CODE);

        curl_close ($curl);

        
        if ($jobtitle == 'null')
        {
            $jobtitle = '';
        }

        if ($company == 'null')
        {
            $company = '';
        }
        

        if ($status_code == 200)
        {
            //var_dump($result);

            foreach (json_decode($result)->entities as $entity)
            {
                if ($entity->type == 'Company')
                {
                    $company = array_pop(explode(' at ', $entity->text));
                    if ($jobtitle == $company)
                    {
                        $jobtitle = '';
                    }
                }

                if ($entity->type == 'JobTitle')
                {
                    $jobtitle = array_shift(explode(' at ', $entity->text));
                    if ($company == $jobtitle)
                    {
                        $company = '';
                    }
                }
            }

            return array("title" => $jobtitle, "organization" => $company);
        }
        else
        {
            //var_dump($result);

            return array("title" => $jobtitle, "organization" => $company);
        }

    }
}

class Slack
{
    private $apikey;
    private $channel;
    private $context;
    private $starttime;

    function __construct()
    {
        $this->apikey = "";
        $this->channel = "";
    }

    function upload($file, $title)
    {
        $title = urlencode($title);

        $comment = urlencode("Here you go! I did my best to fill everything out.");


        $multipart_boundary = '--------------------------' . microtime(true);
        
        $content = "--$multipart_boundary\r\n" . 
            "Content-Disposition: form-data; name=\"file\"; filename=\"" . basename($file) . "\"\r\n" .
            "Content-Type: application/zip\r\n\r\n" . 
            file_get_contents($file) . "\r\n" . 
            "--$multipart_boundary--\r\n";

        $opts = array(
            'http'=>array(
            'method'=>"POST",
            'header'=>"Content-Type: multipart/form-data; boundary=$multipart_boundary",
            'content'=>$content
            )
        );

        $this->context = stream_context_create($opts);
        
        $stuff = file_get_contents("https://slack.com/api/files.upload?token=$this->apikey&channels=$this->channel&filetype=zip&title=$title&initial_comment=$comment", FALSE, $this->context);

        return json_decode($stuff);
    }

    function respond_to_thank()
    {
        $history = json_decode(file_get_contents("https://slack.com/api/groups.history?token=$this->apikey&channel=$this->channel&unreads=true"));

        foreach ($history->messages as $message)
        {
            $messagetime = $message->ts;
            $currenttime = microtime(TRUE);

            if ($currenttime < ($messagetime + 300))
            {
                if (strpos(strtolower($message->text), 'thank') && strpos($message->text, '<@U9TLU0JKD>'))
                {
                    $text = urlencode("You're welcome.");
                    file_get_contents("https://slack.com/api/chat.postMessage?token=$this->apikey&channel=$this->channel&text=$text&as_user=true");
                }
            }
        }
    }
}

$mailchimp = new MailChimp;
$watson = new Watson;
$slack = new Slack;

$time = new Datetime('January 1');
$sincetime = $time->format('Y-m-d\TH:i:sO');

$time = new Datetime('-2 days');
$beforetime = $time->format('Y-m-d\TH:i:sO');

$slack->respond_to_thank();

$reports = $mailchimp->query("reports?count=10&since_send_time=$sincetime&before_send_time=$beforetime")->reports;

var_dump($reports);

foreach ($reports as $report)
{
    $data = array();

    if (!is_dir("$mailchimp->basedir/$report->id"))
    {
        mkdir("$mailchimp->basedir/$report->id");
        $recipients = $mailchimp->query("reports/$report->id/sent-to?count=1000")->sent_to;
        $openers = $mailchimp->query("reports/$report->id/open-details?count=1000")->members;

        $openeremails = array();

        foreach ($openers as $opener)
        {
            $openeremails[] = $opener->email_address;
        }

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Media List');


        echo "recpients for $report->id:\r\n";
        var_dump($recipients);

        foreach ($recipients as $recipient)
        {

            $firstname = $recipient->merge_fields->FNAME;
            $lastname = $recipient->merge_fields->LNAME;

            if (!$recipient->merge_fields->MMERGE5)
            {
                $title = $recipient->merge_fields->MMERGE3;
            }
            else
            {
                $title = $recipient->merge_fields->MMERGE5;
            }

            if (!$recipient->merge_fields->MMERGE6)
            {
                $organization = $recipient->merge_fields->MMERGE4;
            }
            else
            {
                $organization = $recipient->merge_fields->MMERGE6;
            }

            $analysis = $watson->analyze($firstname, $lastname, $title, $organization);

            //var_dump($analysis);

            
            $title = $analysis['title'];
            $organization = $analysis['organization'];

            $email = $recipient->email_address;


            $data[] = array(
                "First Name" => $firstname,
                "Last Name" => $lastname,
                "Title" => $title,
                "Organization" => $organization,
                "Email" => $email
            );

            var_dump($report->id);
            var_dump($data);
        }
        
        $headers = array();

        foreach ($data as $row)
        {
            foreach (array_keys($row) as $header)
            {
                if (!in_array($header, $headers))
                {
                    $headers[] = $header;
                }
            }
        }

        foreach ($data as $rowindex => $row)
        {
            $rowindex++;

            foreach ($row as $cellindex => $cell)
            {
                $letter = chr(65 + array_search($cellindex, $headers));
                $sheet->setCellValue("$letter$rowindex", $cell);

                if (in_array($cell, $openeremails))
                {
                    $firstcell = "A$rowindex";
                    $lastcell = chr(64 + count($headers)) . $rowindex;

                    $sheet->getStyle("$firstcell:$lastcell")->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFFFFF00');
                }
            }
        }


        $sheet->insertNewRowBefore(1, 1);

        foreach ($headers as $index => $header)
        {
            $letter = chr(65 + array_search($header, $headers));
            $cell = $letter . '1';
            $sheet->getStyle($cell)->getFont()->setBold(TRUE);
            $sheet->setCellValue($cell, $header);
            $sheet->getColumnDimension($letter)->setAutoSize(TRUE);
        }

        $filename = preg_replace("/[^A-Za-z0-9_.]/i", "", $report->campaign_title);
                
        $writer = new Xlsx($spreadsheet);
        $writer->save("$mailchimp->basedir/$report->id/$filename.xlsx");

        

        
        $html = $mailchimp->query("campaigns/$report->id/content")->archive_html;
        $mpdf = new \Mpdf\Mpdf();
        $mpdf->WriteHTML($html);
        $mpdf->Output("$mailchimp->basedir/$report->id/$filename.pdf", \Mpdf\Output\Destination::FILE);

        $zip = new ZipArchive();
        $zip->open("$mailchimp->basedir/$report->id/$filename.zip", ZipArchive::CREATE);
        $zip->addFile("$mailchimp->basedir/$report->id/$filename.pdf", "$filename.pdf");
        $zip->addFile("$mailchimp->basedir/$report->id/$filename.xlsx", "$filename.xlsx");
        $zip->close();

        var_dump($slack->upload("$mailchimp->basedir/$report->id/$filename.zip", $report->campaign_title));
    }
}


?>
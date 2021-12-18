<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$file = 'Division Data Set.xlsx';

if (!file_exists($file)) {
    exit("Your file doesn't exist! check & try again.");
}

$spreadsheet = $reader->load($file); // load sheet

$sheet_names = $spreadsheet->getSheetNames(); // all sheet name

$index1 = array_search('TestDataSets', $sheet_names);
if ($index1 === false) {
    exit('TestDataSets.xlsx sheet not found! check & try again.');
}

$inputData = $spreadsheet->getSheet($index1)->toArray(); // get input dataset

$inputData[0][4] = 'Is Matched'; // add new column

foreach ($inputData as $key => $input) {
    $lang = $input[1];
    $lng = $input[2];
    $division = $input[3];

    if ($key > 0) { // except column header
        $input[4] = 'no';

        $index2 = array_search($division, $sheet_names);
        if ($index2 === false) {
            echo $division . ' sheet missing! <br>';
            continue;
        }

        $dataSet = $spreadsheet->getSheet($index2)->toArray(); // get boundary dataset
        unset($dataSet[0]); // remove column header

        foreach ($dataSet as $data) {
            if ($lang == $data[0] && $lng == $data[1]) { // check if match
                $input[4] = 'yes';
            }
        }
    }

    echo $input[0] . "---" . $input[1] . "---" . $input[2] . "---" . $input[3] . "---" . $input[4] . " <br>";
}
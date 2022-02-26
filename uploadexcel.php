<?php
//include the following 2 files
require 'PHPExcel.php';
require_once 'PHPExcel/IOFactory.php';
if(isset($_POST['Submit'])) {
    $SERVER = 'localhost';
    $USERNAME = 'root';
    $PASSWORD = 'Leo28032000leo';
    $DB = 'sotbit';
    $connection = new mysqli($SERVER, $USERNAME, $PASSWORD, $DB);

    $path = "parser.xlsx";

    $objPHPExcel = PHPExcel_IOFactory::load($path);
    foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
        $worksheetTitle = $worksheet->getTitle();
        $highestRow = $worksheet->getHighestRow(); // e.g. 10
        $highestColumn = $worksheet->getHighestColumn(); // e.g 'F'
        $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);
        $nrColumns = ord($highestColumn) - 64;

        echo '<table border="1"><tr>';
        for ($row = 1; $row <= $highestRow; ++$row) {
            echo '<tr>';
            for ($col = 0; $col < $highestColumnIndex; ++$col) {
                $cell = $worksheet->getCellByColumnAndRow($col, $row);
                $val = $cell->getValue();
                $dataType = PHPExcel_Cell_DataType::dataTypeForValue($val);
                echo '<td>' . $val . '</td>';
            }
            echo '</tr>';
        }
        echo '</table>';
    }

    for ($row = 2; $row <= $highestRow; ++$row) {
        $val = array();
        for ($col = 0; $col < $highestColumnIndex; ++$col) {
            $cell = $worksheet->getCellByColumnAndRow($col, $row);
            $val[] = $cell->getValue();
        }

        $Connection = "INSERT INTO sotbit.test(article, name, price, remains) VALUES ('" . $val[1] . "','" . $val[2] . "','" . $val[3] . "','" . $val[4] . "')";
        $connection->query($Connection);

    }
}

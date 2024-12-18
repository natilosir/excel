<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

class Excel
{
    // Convert an array to an Excel file
    public static function ToExcel($data, $filename = "output.xlsx")
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        foreach ($data as $rowIndex => $row) {
            foreach ($row as $colIndex => $value) {
                $cellAddress = self::columnIndexToLetter($colIndex + 1) . ($rowIndex + 1);
                $sheet->setCellValue($cellAddress, $value);
            }
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save($filename);

        echo "File created: $filename\n";
    }

    // Convert an Excel file to an array
    public static function ToArray($filename)
    {
        if (!file_exists($filename)) {
            throw new \Exception("File not found: $filename");
        }

        $spreadsheet = IOFactory::load($filename);
        $sheet = $spreadsheet->getActiveSheet();
        $data = [];

        foreach ($sheet->getRowIterator() as $row) {
            $rowIndex = $row->getRowIndex();
            $rowData = [];

            foreach ($row->getCellIterator() as $cell) {
                $rowData[] = $cell->getValue();
            }

            $data[] = $rowData;
        }

        return $data;
    }

    // Helper to convert column index to letter (e.g., 1 -> A, 2 -> B, etc.)
    private static function columnIndexToLetter($colIndex)
    {
        $letter = '';
        while ($colIndex > 0) {
            $colIndex--;
            $letter = chr(65 + ($colIndex % 26)) . $letter;
            $colIndex = (int)($colIndex / 26);
        }
        return $letter;
    }
}

// Example usage

// Example 1: Create Excel file from array
$data = [
    ["Name", "Age", "Salary"],
    ["John Doe", 28, 5000],
    ["Jane Smith", 34, 7000],
    ["Mark Taylor", 45, 9000],
];
Excel::ToExcel($data, "example.xlsx");

// Example 2: Convert Excel file to array
    $arrayFromExcel = Excel::ToArray("example.xlsx");
    print_r($arrayFromExcel);

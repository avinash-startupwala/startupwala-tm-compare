<?php
  error_reporting(E_ALL);
       ini_set('display_errors', 1);

ini_set('memory_limit', '-1');
error_reporting(0);

function download_excel_file()
{

  unlink("mca_file.xlsx");

  file_put_contents("mca_file.xlsx", fopen('http://www.mca.gov.in/mcafoportal/companiesRegReport.do', 'r'));
  
}

function parse_excel_sheet_and_get_new_cins() 
{
  $array_of_cins = [];
  $excelReader = PHPExcel_IOFactory::createReaderForFile("mca_file.xlsx");
  $excelObj = $excelReader->load("mca_file.xlsx");
  $worksheet = $excelObj->getSheet(0);
  $lastRow = $worksheet->getHighestRow();

  for ($row = 3; $row <= $lastRow; $row++) {
        $the_cin = $worksheet->getCell('A'.$row)->getValue()->getPlainText();
        array_push($array_of_cins, $the_cin);
  }
  return $array_of_cins;
}

function parse_excel_sheet_and_get_new_llpins() 
{
  $array_of_cins = [];
  $excelReader = PHPExcel_IOFactory::createReaderForFile("mca_file.xlsx");
  $excelObj = $excelReader->load("mca_file.xlsx");
  $worksheet = $excelObj->getSheet(2);
  $lastRow = $worksheet->getHighestRow();

  for ($row = 3; $row <= $lastRow; $row++) {
        $the_cin = $worksheet->getCell('A'.$row)->getValue()->getPlainText();
        array_push($array_of_cins, $the_cin);
  }
  return $array_of_cins;
}

function pull_compnay_data_if_not_avaiable($array_of_company_ids, $company_type, &$conn)
{
  foreach ($array_of_company_ids as $cin) {
    $company = new MCACompany($cin, $company_type);
    if ($company->do_we_have_full_company_data($conn) == false)
    {
      $company->save_company_to_db($conn);
      echo "Saving data for company with ID: $cin";
    }
  }
}

download_excel_file();

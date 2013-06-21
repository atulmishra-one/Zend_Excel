<?php
class IndexController extends Zend_Controller_Action
{

    public function init()
    {
        /* Initialize action controller here */
    }
	
	public function generateAction()
	{
		$modelUsers = new Application_Model_Users(); // import model data here
		
		$data = $modelUsers->fetchAll(); // get values from table 
		
		require '../library/PHPExcel/Classes/PHPExcel.php';
		$objPHPExcel = new PHPExcel();
		
		// Set document properties
		$objPHPExcel->getProperties()->setCreator("Atul")
									 ->setLastModifiedBy("atul")
									 ->setTitle("Office 2010 XLSX Test Document")
									 ->setSubject("Office 2010 XLSX Test Document")
									 ->setDescription("Test document for Office 2010 XLSX.")
									 ->setKeywords("office 2007 openxml php")
									 ->setCategory("Test result file");
									 
		
		// Add some header
		$objPHPExcel->setActiveSheetIndex(0)
            		->setCellValue('A1', 'Id')
           			->setCellValue('B1', 'Username')
           			->setCellValue('C1', 'Password');

		// Rows
		
		$i = 2;
		
		foreach( $data as $value ) :
		
		$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A'.$i, $value->users_id)
			->setCellValue('B'.$i, $value->username)
			->setCellValue('C'.$i, $value->password);
			
		 $i++;
		
		endforeach;
		// Rename worksheet
		$objPHPExcel->getActiveSheet()->setTitle('Simple');


		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(0);
		
		
		 $this->getResponse()->setRawHeader( "Content-Type: application/vnd.ms-excel; charset=UTF-8")
        ->setRawHeader("Content-Disposition: attachment; filename=excel.xls")
        ->setRawHeader("Content-Transfer-Encoding: binary")
        ->setRawHeader("Expires: 0")
        ->setRawHeader("Cache-Control: must-revalidate, post-check=0, pre-check=0")
        ->setRawHeader("Pragma: public")
        ->setRawHeader("Content-Length: " . strlen($objPHPExcel))
        ->sendResponse();

		$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		$objWriter->save('php://output');
		exit();
	}

}


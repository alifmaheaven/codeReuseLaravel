# phpspredsheet




# code reuse for phpspredsheet in laravel

 

### Installation

Pertama install phpspreedshetnya di terminal

```
$ composer require phpoffice/phpspreadsheet
```

#### dicontrollernya jadikan begini

```sh
<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
//use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//use PhpOffice\PhpSpreadsheet\Spreadsheet;
 use Symfony\Component\HttpFoundation\StreamedResponse;
use Symfony\Component\Routing\Annotation\Route;
// use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use PhpOffice\PhpSpreadsheet\Writer as Writer;
//styling
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;



class EksportController extends Controller
{
    public function halo()
    {
        $excel = new Spreadsheet();
        $excel->getProperties()->setCreator('My Notes Code')                 
        ->setLastModifiedBy('My Notes Code')                 
        ->setTitle("DBMS")                 
        ->setSubject("DBMS")                 
        ->setDescription("DBMS Customer")                 
        ->setKeywords("Customer");

        $excel->setActiveSheetIndex(0)->setCellValue('A1', "Customer"); // Set kolom A1 dengan tulisan "DATA SISWA"    
        $excel->getActiveSheet()->mergeCells('A1:B1'); // Set Merge Cell pada kolom A1 sampai E1    
        //$excel->getActiveSheet()->getStyle('A1')->getFont()->setBold(TRUE); // Set bold kolom A1    
        //$excel->getActiveSheet()->getStyle('A1')->getFont()->setSize(15); // Set font size 15 untuk kolom A1    
        //$excel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER); // Set text center untuk kolom A1  
        
        //background of cells
        // $excel->getActiveSheet()->getStyle('A1:C3')->getFill()->applyFromArray( [ 
        //     'fillType' => Fill::FILL_SOLID,
        //     //'rotation' => 0, 
        //     'startColor' => [ 
        //         'rgb' => '000000' 
        //     ], 
        //     // 'endColor' => [ 
        //     //     'argb' => 'FFFFFFFF' 
        //     //     ] 
        //         ] 
        //     ); 


        $styleArray = array(
            'fill' => array(
                'fillType' => Fill::FILL_SOLID,
                'color' => array('argb' => '00000000'),
            ),
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => Border::BORDER_THICK,
                    'color' => array('argb' => 'FFFF0000'),
                ),
            ),
            
            'font'  => array(
                        'bold'  => true,
                        'size'  => 15,
                       )
        );
        
        $excel->getActiveSheet()->getStyle('B2:G8')->applyFromArray($styleArray);

      


        $writer = new Writer\Xlsx($excel);

        $response =  new StreamedResponse(
            function () use ($writer) {
                $writer->save('php://output');
            }
        );
        
        $response->headers->set('Content-Type', 'application/vnd.ms-excel');
        $response->headers->set('Content-Disposition', 'attachment;filename="ExportScan.xlsx"');
        $response->headers->set('Cache-Control','max-age=0');
        return $response;
    }
}


```


<?php

namespace App\Controller;

use App\Service\ChunkReadFilterService;
use App\Service\ReaderFilter;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Exception;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use Symfony\Bundle\FrameworkBundle\Controller\AbstractController;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\Routing\Annotation\Route;


class HomeController extends AbstractController
{
    /**
     * @Route("/", name="home")
     */
    public function index()
    {

        return $this->render('home/index.html.twig', [

        ]);
    }

    /**
     * @Route("/excel", name="excel")
     *
     * @return Response
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function excelReader(){

        $reader = new Xlsx();
        $reader->setReadDataOnly(TRUE);
        $spreadsheet = $reader->load("uploads/test.xlsx");
        $worksheet = $spreadsheet->getActiveSheet();
        $sheetdata = $worksheet->toArray(null, true, true, true);

        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
        $res = array();
        for($row=1; $row < $highestRow ; $row++){
            for($col = 1; $col <= $highestColumnIndex; $col++){
                $value = $worksheet->getCellByColumnAndRow($col,$row)->getValue();
                array_push($res,$value);
            }
        }


        return $this->render('home/excel.html.twig', [
            'lists' => $sheetdata,

        ]);

    }

    /**
     * @Route("/filter", name="filter")
     * @return Response
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws Exception
     */
    public function filter():Response
    {
        $reader = new Xlsx();
        $inputFileName = 'uploads/test.xlsx';
        $inputFileType = 'Xlsx';
        $sheetname = 'POS_TRANSACTIONS';
        $filterSubset = new ReaderFilter(1, 10, range('A', 'M'));

        $reader = IOFactory::createReader($inputFileType);

        $reader->setLoadSheetsOnly($sheetname);


        $reader->setReadFilter($filterSubset);
        $spreadsheet = $reader->load($inputFileName);
        $worksheet = $spreadsheet->getActiveSheet();

        $sheetData = $worksheet->toArray(null, true, true, true);

        return $this->render('home/excel.html.twig', [
            'lists' => $sheetData,

        ]);

    }

    /**
     * @Route("/filters", name="special_filter")
     * @return Response
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function specialFilter():Response
    {
        $reader = new Xlsx();
        $reader->setReadDataOnly(TRUE);
        $spreadsheet = $reader->load("uploads/test.xlsx");
        $worksheet = $spreadsheet->getActiveSheet();
        $sheetdata = $worksheet->toArray(null, true, true, true);


        $worksheet->fromArray( null, NULL, 'A1' )
            ->fromArray( $sheetdata, NULL, 'A4' );
        $result = $worksheet->setCellValue('A12', '=DCOUNT(A4:E10,"Height",A1:B3)');

        return $this->render('home/excel.html.twig', [
            'list' => $result,
        ]);
    }

    /**
     * @Route("/chunkfilter", name="chunkfilter")
     * @return Response
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function chunkFilter():Response
    {
        $inputFileType = 'Xlsx' ;
        $inputFileName = 'uploads/test.xlsx' ;
        $reader = IOFactory::createReader($inputFileType);

        /**  Define how many rows we want to read for each "chunk"  **/
        $chunkSize = 2;
        /**  Create a new Instance of our Read Filter  **/
        $chunkFilter = new ChunkReadFilterService();


        /**  Loop to read our worksheet in "chunk size" blocks  **/
        for ($startRow = 2; $startRow <= 240; $startRow += $chunkSize) {
            /**  Tell the Read Filter which rows we want this iteration  **/
            $chunkFilter->setRows($startRow, $chunkSize);
            /**  Tell the Reader that we want to use the Read Filter  **/
            $reader->setReadFilter($chunkFilter);
            /**  Load only the rows that match our filter  **/
            $spreadsheet = $reader->load($inputFileName);
            //    Do some processing here
            $sheetData = $spreadsheet->getActiveSheet()->toArray(null, true, true, true);
        }



        return $this->render('home/filter.html.twig', [
            'list' => $sheetData,
        ]);
    }
}

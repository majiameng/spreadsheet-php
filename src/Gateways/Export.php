<?php
/**
 * @name: 报表导出查询
 * @Created by IntelliJ IDEA
 * @author: tinymeng
 * @file: Export.php
 * @Date: 2018/7/4 10:15
 */
namespace tinymeng\spreadsheet\Gateways;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Exception as ExceptionAlias;
use tinymeng\spreadsheet\Connector\Gateway;
use tinymeng\tools\File;
use tinymeng\spreadsheet\Excel\TWorkSheet;
use tinymeng\spreadsheet\Util\TConfig;

class Export extends Gateway {

    /**
     * 获取sheet表格数目
     * @var
     */
    private $sheetCount = 1;


    /**
     * TCfonig
     */
    use TConfig;

    /**
     * TWorkSheet
     */
    use TWorkSheet;

    /**
     * __construct
     */
    public function __construct($config=[]){
        $this->setConfig($config);

        $this->spreadSheet = new Spreadsheet();
        //初始化表格格式
        $this->initSpreadSheet();
        return $this;
    }

    /**
     * @param $config
     * @return $this
     */
    public function setConfig($config){
        foreach ($config as $key => $value) {
            if (property_exists($this, $key)) {
                $this->$key = $value;
            }
        }
        return $this;
    }

    /**
     * 初始化表格格式
     * initSpreadSheet
     * @return void
     */
    public function initSpreadSheet()
    {
        /** 实例化定义默认excel **/
        $this->spreadSheet->getProperties()->setCreator($this->creator)->setLastModifiedBy($this->creator);
        if($this->horizontalCenter){
            $this->spreadSheet->getDefaultStyle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER); //默认水平居中
            $this->spreadSheet->getDefaultStyle()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER); //默认垂直居中
            $this->spreadSheet->getDefaultStyle()->getAlignment()->setHorizontal(Alignment::VERTICAL_CENTER); //默认垂直居中
        }
    }

    /**
     * 创建新的sheet
     * @param $sheetName
     * @return $this
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function createWorkSheet($sheetName="Worksheet")
    {
        $this->sheetName = $sheetName;
        /** @var workSheet */
        $this->workSheet = $this->greateWorkSheet($sheetName);
        if($this->workSheet == null){
            if($this->sheetCount==1){
                $this->workSheet = $this->spreadSheet->getActiveSheet();
            }else{
                $this->workSheet = $this->spreadSheet->createSheet();
            }
            $this->sheetCount += 1;//总sheet数量
            $this->workSheet->setTitle($sheetName);//设置sheet名称
        }

        /** 初始化当前workSheet */
        $this->initWorkSheet();

        return $this;
    }

    /**
     * @param $sheetName
     * @return workSheet
     */
    public function greateWorkSheet($sheetName)
    {
        /** @var workSheet */
        return $this->spreadSheet->getSheetByName($sheetName);
    }

    /**
     * 生成excel表格
     * @return $this
     */
    public function generate()
    {
        /** 开启自动筛选 **/
        if($this->autoFilter){
            $this->spreadSheet->getActiveSheet()->setAutoFilter(
                $this->spreadSheet->getActiveSheet()->calculateWorksheetDimension()
            );
        }
        //文件存储
        if(empty($this->fileName)){
            $this->getFileName($this->sheetName);
        }
        return $this;
    }


    /**
     * @param $file_name
     * @return string
     */
    private function getFileName($sheetName){
        $this->fileName = $sheetName.'_'.date('Y-m-d').'_'.rand(111,999).'.xlsx';
        return $this->fileName;
    }

    /**
     * 文件下载
     * @param $filename
     * @return void
     * @throws ExceptionAlias
     */
    public function download($filename){
        if(empty($filename)){
            $filename = $this->fileName;
        }else{
            $filename = $this->getFileName($filename);
        }

        /** 输出下载 **/
        ob_end_clean();//清除缓冲区,避免乱码
        header( 'Access-Control-Allow-Headers:responsetype,content-type,usertoken');
        header( 'Access-Control-Allow-Methods:GET,HEAD,PUT,POST,DELETE,PATCH');
        header( 'Access-Control-Allow-Origin:*');
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$filename);
        header('Cache-Control: max-age=0');

        $objWrite = IOFactory::createWriter($this->spreadSheet, 'Xlsx');
        $objWrite->save('php://output');
        exit();
    }

    /**
     * 文件存储
     * @param $filename
     * @param $pathName
     * @return string
     * @throws ExceptionAlias
     */
    public function save($filename='',$pathName=''): string
    {
        $pathName = $this->getPathName($pathName);
        if(empty($filename)){
            $filename = $this->fileName;
        }else{
            $filename = $this->getFileName($filename);
        }
        File::mkdir($pathName);
        $objWrite = IOFactory::createWriter($this->spreadSheet, 'Xlsx');
        $objWrite->save($pathName.$filename);
        return $pathName.$filename;
    }

}

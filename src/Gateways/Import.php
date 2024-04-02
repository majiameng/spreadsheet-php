<?php
/**
 * @name: 报表导入查询
 * @Created by IntelliJ IDEA
 * @author: tinymeng
 * @file: Import.php
 * @Date: 2018/7/4 10:15
 */
namespace tinymeng\spreadsheet\Gateways;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use tinymeng\spreadsheet\Connector\Gateway;
use tinymeng\spreadsheet\Excel\SpreadSheet;
use tinymeng\spreadsheet\Util\TConfig;

class Import extends Gateway {

    /**
     * TConfig
     */
    use TConfig;
    /**
     * TSpreadSheet
     */
    use SpreadSheet;


    /**
     * @return $this
     * @throws Exception
     */
    public function initWorkSheet($filename='')
    {
        if(!empty($filename)){
            $this->setFileName($filename);
        }
        $this->spreadSheet = IOFactory::load($this->fileName);
        //获取sheet表格数目
        $this->sheetCount = $this->spreadSheet->getSheetCount();
        //默认选中sheet0表
        $this->spreadSheet->setActiveSheetIndex($this->sheet);
        $this->workSheet = $this->spreadSheet->getActiveSheet();
        //获取表格行数
        $this->rowCount = $this->workSheet->getHighestDataRow();
        //获取表格列数
        $this->columnCount = $this->workSheet->getHighestDataColumn();
        //初始化所有列数组
        $this->cellName = $this->getCellName($this->columnCount);
        return $this;
    }


}

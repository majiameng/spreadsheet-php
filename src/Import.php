<?php
/**
 * @name: 报表导入查询
 * @Created by IntelliJ IDEA
 * @author: tinymeng
 * @file: Import.php
 * @Date: 2018/7/4 10:15
 */
namespace tinymeng\spreadsheet;

use PhpOffice\PhpSpreadsheet\IOFactory;
use tinymeng\spreadsheet\Connector\Gateway;

class Import extends Gateway {

    /**
     * 表格的sheet
     * @var int
     */
    public $sheet = 0;
    /**
     * 获取sheet表格数目
     * @var
     */
    public $sheetCount;
    /**
     * 表头所在行
     * @var int
     */
    public $titleFieldsRow = 1;

    /**
     * 获取表格列数
     * @var
     */
    public $columnCount;

    /**
     * 获取表格行数
     * @var
     */
    public $rowCount;
    /**
     * title
     * @var
     */
    public $title;
    /**
     * title字段
     * @var
     */
    public $title_fields;
    /**
     * 文件名称
     * @var
     */
    public $fileName;
    private $cellName = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ'];

    /**
     * initReadExcel
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:31
     */
    public function initReadExcel(){
        $this->spreadSheet = $spreadsheet = IOFactory::load($this->fileName);
        //获取sheet表格数目
        $this->sheetCount = $spreadsheet->getSheetCount();
        //默认选中sheet0表
        $this->spreadSheet->setActiveSheetIndex($this->sheet);
        $this->workSheet = $this->spreadSheet->getActiveSheet();
        //获取表格行数
        $this->rowCount = $this->workSheet->getHighestDataRow();
        //获取表格列数
        $this->columnCount = $this->workSheet->getHighestDataColumn();
    }

    public function setTitle($title){
        $this->title = $title;
    }

    /**
     * getExcelData
     * @param $fields
     * @return array
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:30
     */
    public function getExcelData($fields){
        /* 循环读取每个单元格的数据 */
        $result = [];
        $cellName = array_slice($this->cellName,0,count($this->title));
        $dataRow = $this->titleFieldsRow+1;
        //行数循环
        for ($row = $dataRow; $row <= $this->rowCount; $row++){
            $rowFlog = false;//行是否有内容（过滤空行）
            //列数循环 , 列数是以A列开始
            $data = [];
            foreach ($cellName as $column){
                $value = trim($this->workSheet->getCell($column.$row)->getValue());
                if(isset($fields[$column])){
                    $data[$fields[$column]] = $value;
                    if(!empty($value)) $rowFlog = true;//有内容
                }
            }
            if($rowFlog) $result[] = $data;
        }
        return $result;
    }

    /**
     * getTitle
     * @return mixed
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:30
     */
    public function getTitle(){
        return $this->title;
    }

    /**
     * getTitleFields
     * @return array
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:30
     */
    public function getTitleFields(){
        $title = $this->getTitle();

        $row = $this->titleFieldsRow;
        $titleDataArr = [];
        $cellName = array_slice($this->cellName,0,count($title));

        foreach ($cellName as $column){
            $value = trim($this->workSheet->getCell($column.$row)->getValue());
            if(!empty($value)){
                $titleDataArr[$value] = $column;
            }
        }
        $title_fields = [];
        foreach ($title as $key=>$value) {
            if(isset($titleDataArr[$key])){
                $title_fields[$titleDataArr[$key]] = $value;
            }
        }
        $this->title_fields = $title_fields;
        return $title_fields;
    }
}

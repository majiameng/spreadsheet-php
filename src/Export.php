<?php
/**
 * @name: 报表导出查询
 * @Created by IntelliJ IDEA
 * @author: tinymeng
 * @file: Export.php
 * @Date: 2018/7/4 10:15
 */
namespace tinymeng\spreadsheet;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Writer\Exception as ExceptionAlias;
use tinymeng\spreadsheet\Connector\Gateway;
use tinymeng\tools\exception\StatusCode;
use tinymeng\tools\exception\TinymengException;
use tinymeng\tools\File;

class Export extends Gateway {

    /**
     * sheet名称
     * @var
     */
    public $sheetName;
    /**
     * 文件名称
     * @var
     */
    public $fileName;
    /**
     * 文件名称
     * @var
     */
    public $group_left;
    /**
     * 查询数据
     * @var
     */
    public $data;
    /**
     * 报表名称(主标题)
     * @var
     */
    public $mainTitle;
    /**
     * 是否需要报表名称(主标题)
     * @var bool
     */
    public $mainTitleLine = false;
    /**
     * 存储方式: download下载, save存储本地
     * @var string
     */
    public $saveType = 'download';
    /**
     * 定义表头行高
     * @var int 常用：22
     */
    public $titleHeight = null;
    /**
     * 定义表头列宽(未设置则自动计算宽度)
     * @var int 常用：20
     */
    public $titleWidth = null;
    /**
     * 定义数据行高
     * @var int 常用：22
     */
    public $height = null;

    /**
     * 自动筛选(是否开启)
     * @var bool
     */
    public $autoFilter = false;
    /**
     * 是否居中
     * @var string
     */
    public $horizontal_center = true;
    /**
     * 自动适应文本类型
     * @var bool
     */
    public $autoDataType = true;

    /**
     * 冻结窗格（要冻结的首行首列"B2"，false不开启）
     * @var string|bool
     */
    public $freezePane = false;

    /**
     * 文件信息
     * @var array
     */
    public $fileTitle=[];

    /**
     * 导出文件路径名称
     * @var string
     */
    public $pathName;

    /**
     * 定义默认列数
     * @var int
     */
    private $_col = 0;
    /**
     * 定义当前行数
     * @var int
     */
    private $_row = 1;
    /**
     * 定义所有字段
     * @var array
     */
    private $field = [];

    /**
     * excel生成并下载
     * @return mixed
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:31
     */
    public function exportExcel()
    {
        $this->fileTitle['title_row'] = $this->fileTitle['title_row'] ?? 1;          //标题占用行数
        $this->group_left = $this->fileTitle['group_left'] ?? [];      //左侧分组字段

        /** 实例化定义默认excel **/
        $this->spreadSheet = new Spreadsheet();
        $this->spreadSheet->getProperties()->setCreator("TinyMeng")->setLastModifiedBy("TinyMeng");
        if($this->horizontal_center){
            $this->spreadSheet->getDefaultStyle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER); //默认水平居中
            $this->spreadSheet->getDefaultStyle()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER); //默认垂直居中
            $this->spreadSheet->getDefaultStyle()->getAlignment()->setHorizontal(Alignment::VERTICAL_CENTER); //默认垂直居中
        }

        $this->workSheet = $this->spreadSheet->getActiveSheet();
        if($this->freezePane) $this->workSheet->freezePane($this->freezePane); //冻结窗格
        $this->workSheet->setTitle($this->sheetName);   //设置sheet名称

        /** 设置表头 **/
        $this->excelTitle();
        /** 设置第一行格式 */
        $this->excelHeader();

        /** 获取列表里所有字段 **/
        foreach ($this->fileTitle['title'] as $key => $val){
            if(is_array($val)){
                foreach ($val as $k => $v){
                    $this->field[] = $v;
                }
            }else{
                $this->field[] = $val;
            }
        }
        /** 查询结果赋值 **/
        if(!empty($this->data)){
            $this->excelSetValue();
        }
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
        $saveType = $this->saveType;
        $result = $this->$saveType();
        return $result;
    }

    /**
     * excelSetValue
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:43
     */
    public function excelSetValue(){
        if(empty($this->group_left)){ //判断左侧是否分组
            foreach ($this->data as $key => $val){
                $this->excelSetCellValue($val);
            }
        }else{   //根据设置分组字段进行分组
            /** 数据分组 **/
            $data = [];
            $group_left_count = count($this->group_left);
            if($group_left_count == 1){
                foreach ($this->data as $k => $v){
                    $data[$v[$this->group_left[0]]][] = $v;
                }
                foreach ($data as $k =>$v){
                    $data[$k] = [
                        'data' => $v,
                        'count' => count($v)
                    ];
                }
                $this->excelGroupLeft($data, 0, $group_left_count);
            }elseif($group_left_count == 2){
                foreach ($this->data as $v) {
                    $data[$v[$this->group_left[0]]][$v[$this->group_left[1]]][] = $v;
                }
                $this->data = $this->arrayCount($data);
                $this->excelGroupLeft($this->data, 0, $group_left_count);
            }else{
                throw new TinymengException(StatusCode::COMMON_PARAM_INVALID,
                    '左侧分组过多，导出失败！'
                );
            }
        }
    }

    /**
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function excelHeader(){
        if($this->mainTitleLine == true){
            $row = 1;
            $this->workSheet->setCellValue('A'.$row, $this->mainTitle);
            $this->workSheet->mergeCells('A'.$row.':'.$this->cellName($this->_col-1).$row);
            $this->workSheet->getRowDimension($row)->setRowHeight('25');
        }
    }

    /**
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function excelTitle(){
        if($this->mainTitleLine == true){
            $this->_row ++;//当前行数
        }

        $_merge = $this->cellName($this->_col);
        foreach ($this->fileTitle['title'] as $key => $val) {
            if(!empty($this->titleHeight)) {
                $this->workSheet->getRowDimension($this->_col)->setRowHeight($this->titleHeight);//行高度
            }
            $rowName = $this->cellName($this->_col);
            $this->workSheet->getStyle($rowName . $this->_row)->getAlignment()->setWrapText(true);//自动换行
            if (is_array($val)) {
                $num = 1;
                $_cols = $this->_col;
                foreach ($val as $k => $v) {
                    $this->workSheet->setCellValue($this->cellName($_cols) . ($this->_row+1), $k);
                    if(!empty($this->titleWidth)) {
                        $this->workSheet->getColumnDimension($this->cellName($_cols))->setWidth($this->titleWidth); //列宽度
                    }else{
                        $this->workSheet->getColumnDimension($this->cellName($_cols))->setAutoSize(true); //自动计算宽度
                    }
                    if ($num < count($val)) {
                        $this->_col++;
                        $num++;
                    }
                    $_cols++;
                }
                $this->workSheet->mergeCells($_merge . $this->_row.':' . $this->cellName($this->_col) .$this->_row);
                $this->workSheet->setCellValue($_merge . $this->_row, $key);//设置值
            } else {
                if ($this->fileTitle['title_row'] != 1) {
                    $this->workSheet->mergeCells($rowName . $this->_row.':' . $rowName . ($this->_row + $this->fileTitle['title_row'] - 1));
                }
                $this->workSheet->setCellValue($rowName . $this->_row, $key);//设置值
                if(!empty($this->titleWidth)){
                    $this->workSheet->getColumnDimension($rowName)->setWidth($this->titleWidth); //列宽度
                }else{
                    $this->workSheet->getColumnDimension($rowName)->setAutoSize(true); //自动计算宽度
                }
            }
            $this->_col++;
            $_merge = $this->cellName($this->_col);
        }
        $this->_row += $this->fileTitle['title_row'];//当前行数
    }

    /**
     * excel单元格赋值
     * @author tinymeng
     * @param array $val 数据
     */
    private function excelSetCellValue($val)
    {
        //设置单元格行高
        if(!empty($this->height)){
            $this->workSheet->getRowDimension($this->_row)->setRowHeight($this->height);
        }
        $_lie = 0;
        foreach ($this->field as $v){
            $rowName = $this->cellName($_lie);
            if(strpos($v,'.') !== false){
                $v = explode('.',$v);
                $content = $val;
                for ($i=0;$i<count($v);$i++){
                    $content = $content[$v[$i]]??'';
                }
            }else{
                $content = ($val[$v]??'');
            }
            if(is_array($content) && isset($content['type']) && isset($content['content'])){
                if($content['type'] == 'image'){
                    $path = $this->verifyFile($content['content']);
                    $drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
                    $drawing->setPath($path);
                    if(!empty($content['height'])) {
                        $drawing->setHeight($content['height']);
                    }
                    if(!empty($content['width'])) {
                        $drawing->setWidth($content['width']);//只设置高，宽会自适应，如果设置宽后，高则失效
                    }
                    $drawing->setCoordinates($rowName.$this->_row);
                    $drawing->setWorksheet($this->workSheet);
                }
            }else {
                $content = $this->formatValue($content);//格式化数据
                if (is_numeric($content)){
                    if($this->autoDataType && strlen($content)<11){
                        $this->workSheet->setCellValueExplicit($rowName.$this->_row, $content,DataType::TYPE_NUMERIC);
                    }else{
                        $this->workSheet->setCellValueExplicit($rowName.$this->_row, $content,DataType::TYPE_STRING2);
                    }
                }else{
                    $this->workSheet->setCellValueExplicit($rowName.$this->_row, $content,DataType::TYPE_STRING2);
                }
            }
            $_lie ++;
        }
        $this->_row ++;
    }

    /**
     * 单元格合并并赋值
     * @param array $data 数据
     * @param int $_lie   开始行数
     * @param $group_left_count
     * @author tinymeng
     */
    private function excelGroupLeft(array $data, int $_lie, $group_left_count)
    {
        $group_start = $this->_col; //二级合并单元格开始
        foreach ($data as $key => $val){
            $rowName = $this->cellName($_lie);
            $coordinate = $rowName.$this->_col.':'.$rowName.($this->_col+$val['count']-1);
            $this->workSheet->mergeCells($coordinate);
            if($group_left_count == 1){
                foreach ($val['data'] as $data){
                    $this->excelSetCellValue($data);
                }
            }else{
                $rowName = $this->cellName($_lie+1);  //对应的列值
                foreach ($val['data'] as $k => $v){
                    $group_end_col = $group_start + $v['count']-1;
                    $coordinate = $rowName.$group_start.':'.$rowName.$group_end_col;
                    $this->workSheet->mergeCells($coordinate);
                    $group_start = $group_end_col+1;
                    foreach ($v['data'] as $data){
                        $this->excelSetCellValue($data);
                    }
                }
            }
        }

    }

    /**
     * 二位数组获取每一级别数量
     * @author tinymeng
     * @param array $data 二维数组原始数据
     * @return array
     */
    private function arrayCount($data=[])
    {
        foreach ($data as $key => $val){
            $num = 0;
            foreach ($val as $k => $v){
                $sub_num = count($v);
                $num = $num+$sub_num;
                $val[$k] = [
                    'count' => $sub_num,
                    'data' => $v
                ];
            }
            $data[$key] = [
                'count' => $num,
                'data' => $val
            ];
        }
        return $data;
    }

    /**
     * @param $file_name
     * @return string
     */
    private function getFileName($sheetName){
        $this->fileName = $fileName = $sheetName.'_'.date('Y-m-d').'_'.rand(111,999).'.xlsx';
        return $fileName;
    }

    /**
     * 文件下载
     * @return void
     * @throws ExceptionAlias
     */
    private function download(){
        $filename = $this->fileName;

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
     * @return string
     * @throws ExceptionAlias
     */
    private function save(): string
    {
        //删除当前目录下的同名文件
        $filename = $this->fileName;
        if(empty($this->pathName)) $this->pathName = dirname( dirname(dirname(dirname(dirname(__FILE__))))).DIRECTORY_SEPARATOR."public".DIRECTORY_SEPARATOR."export".DIRECTORY_SEPARATOR.date('Ymd').DIRECTORY_SEPARATOR;
        File::mkdir($this->pathName);
        $objWrite = IOFactory::createWriter($this->spreadSheet, 'Xlsx');
        $objWrite->save($this->pathName.$filename);
        return $this->pathName.$filename;
    }

}

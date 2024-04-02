<?php
/**
 * @name: TWorklSheet
 * @author: JiaMeng <666@majiameng.com>
 * @file: Export.php
 * @Date: 2024/03/04 10:15
 */
namespace tinymeng\spreadsheet\Excel;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use tinymeng\tools\exception\StatusCode;
use tinymeng\tools\exception\TinymengException;

trait TWorkSheet{

    /**
     * sheet名称
     * @var
     */
    private $sheetName;
    /**
     * 查询数据
     * @var
     */
    private $data;
    /**
     * 报表名称(主标题)
     * @var
     */
    private $mainTitle;
    /**
     * 是否需要报表名称(主标题)
     * @var bool
     */
    private $mainTitleLine = false;

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
     * 文件信息
     * @var array
     */
    private $fileTitle=[];

    /**
     * 标题占用行数
     * @var int
     */
    private $title_row = 1;

    /**
     * 左侧分组字段
     * @var array
     */
    private $group_left = [];


    /**
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function initWorkSheet()
    {
        $this->_col = 0;
        $this->_row = 1;
        $this->fileTitle = [];
        $this->data = [];
        $this->field = [];
        if($this->freezePane) $this->workSheet->freezePane($this->freezePane); //冻结窗格
    }

    /**
     * @param $fileTitle
     * @param $data
     * @return $this
     * @throws TinymengException
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function setWorkSheetData($fileTitle,$data)
    {
        if(isset($fileTitle['title_row']) || isset($fileTitle['group_left'])){
            /**
             * $fileTitle = [
             *       'title_row'=>1,
             *       'group_left'=>[],
             *       'title'=>[
             *           '姓名'=>'name'
             *       ],
             *  ];
             */
            $this->title_row = $fileTitle['title_row']??1;
            $this->group_left = $fileTitle['group_left']??[];
            $this->fileTitle = $fileTitle['title']??[];
        }else{
            /**
             *  $fileTitle = [
             *       '姓名'=>'name',
             *  ];
             */
            $this->fileTitle = $fileTitle;
        }
        $this->data = $data;

        /** 设置第一行格式 */
        if($this->mainTitleLine == true){
            $this->excelHeader();
        }

        /** 设置表头 **/
        $this->excelTitle();

        /** 获取列表里所有字段 **/
        foreach ($this->fileTitle as $key => $val){
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
        return $this;
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
    public function excelHeader(){
        $row = 1;
        $this->workSheet->setCellValue('A'.$row, $this->mainTitle);
        $this->workSheet->mergeCells('A'.$row.':'.$this->cellName($this->_col-1).$row);
        $this->workSheet->getRowDimension($row)->setRowHeight('25');
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
        foreach ($this->fileTitle as $key => $val) {
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
                if ($this->title_row != 1) {
                    $this->workSheet->mergeCells($rowName . $this->_row.':' . $rowName . ($this->_row + $this->title_row - 1));
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
        $this->_row += $this->title_row;//当前行数
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
            }elseif($v == '_id'){
                $content = $this->_row-$this->title_row;//自增序号列
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


}

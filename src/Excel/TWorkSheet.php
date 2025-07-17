<?php
/**
 * @name: TWorklSheet
 * @author: JiaMeng <666@majiameng.com>
 * @file: Export.php
 * @Date: 2024/03/04 10:15
 */
namespace tinymeng\spreadsheet\Excel;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use tinymeng\spreadsheet\Util\ConstCode;
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
     * 左侧分组字段
     * @var array
     */
    private $group_left = [];


    /**
     * 获取sheet表格数目
     * @var
     */
    private $sheetCount = 1;

    /**
     * 字段映射方式
     * @var int
     */
    private $fieldMappingMethod = ConstCode::FIELD_MAPPING_METHOD_FIELD_CORRESPONDING_NAME;

    /**
     * 需要自动合并的字段
     * @var array
     */
    private $mergeColumns = [];

    /**
     * 小计行样式
     * @var array
     */
    private $subtotalStyle = [];
    /**
     * 全表样式
     * @var array
     */
    private $sheetStyle = [];

    /**
     * 用户自定义表格操作回调
     * @var callable|null
     */
    private $complexFormatCallback = null;
    /**
     * @var array
     */
    private $titleConfig = [];

    /**
     * @param $data
     * @return $this
     */
    public function setData($data){
        $this->data = $data;
        return $this;
    }

    /**
     * @param $data
     * @return $this
     */
    public function getData(){
        return $this->data;
    }


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
    public function setWorkSheetData($titleConfig,$data)
    {
        $this->titleConfig = $titleConfig;
        if(isset($titleConfig['title_row']) || isset($titleConfig['group_left'])){
            /**
             * $titleConfig = [
             *       'title_row'=>1,
             *       'group_left'=>[],
             *       'title'=>[
             *           '姓名'=>'name'
             *       ],
             *  ];
             */
            $this->title_row = $titleConfig['title_row']??1;
            $this->group_left = $titleConfig['group_left']??[];
            $titleData = $titleConfig['title']??[];
            // 新增：读取mergeColumns配置
            if (isset($titleConfig['mergeColumns'])) {
                $this->mergeColumns = $titleConfig['mergeColumns'];
            }
        }else{
            /**
             *  $titleConfig = [
             *       '姓名'=>'name',
             *  ];
             */
            $titleData = $titleConfig;
        }
        // 根据字段映射方式处理 title
        if ($this->fieldMappingMethod === ConstCode::FIELD_MAPPING_METHOD_FIELD_CORRESPONDING_NAME) {
            $this->fileTitle = array_flip($titleData);// 字段对应名称方式 - 需要将键值对调
        }else{
            $this->fileTitle = $titleData;// 名称对应字段方式 - 保持原样
        }
        $this->data = $data;

        /** 设置第一行格式 */
        if(!empty($this->mainTitle)){
            $this->excelHeader();
            $this->_row ++;//当前行数
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
        // 新增：应用全表样式
        $this->applySheetStyle();
        // 新增：调用自定义表格操作回调
        if (is_callable($this->complexFormatCallback)) {
            call_user_func($this->complexFormatCallback, $this->workSheet);
        }
        return $this;
    }

    /**
     * excelSetValue
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:43
     */
    public function excelSetValue(){
        if(!empty($this->titleConfig['data_start_row'])){
            $this->_row = $this->titleConfig['data_start_row'];
        }

        if(empty($this->group_left)){ //判断左侧是否分组
            $rowStart = $this->_row;
            foreach ($this->data as $key => $val){
                $this->excelSetCellValue($val);
            }
            // 新增：处理mergeColumns自动合并
            if (!empty($this->mergeColumns)) {
                $this->autoMergeColumns($rowStart, $this->_row - 1);
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
                $this->excelGroupLeft($data, $group_left_count);
            }elseif($group_left_count == 2){
                foreach ($this->data as $v) {
                    $data[$v[$this->group_left[0]]][$v[$this->group_left[1]]][] = $v;
                }
                $this->data = $this->arrayCount($data);
                $this->excelGroupLeft($this->data, $group_left_count);
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
        if(!empty($this->mainTitle)){
            $this->workSheet->setCellValue('A'.$row, $this->mainTitle);
        }

        // 计算实际的标题列数
        $titleCount = 0;
        foreach ($this->fileTitle as $val) {
            if (is_array($val)) {
                $titleCount += count($val); // 如果是数组，加上子项的数量
            } else {
                $titleCount++; // 如果是单个标题，加1
            }
        }

        // 使用实际的标题列数来合并单元格
        $this->workSheet->mergeCells('A'.$row.':'.$this->cellName($titleCount-1).$row);
    }

    /**
     * @return void
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    private function excelTitle(){
        if(isset($this->titleConfig['title_show']) && $this->titleConfig['title_show']===false){
            return;
        }
        if(!empty($this->titleConfig['title_start_row'])){
            $this->_row = $this->titleConfig['title_start_row'];
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
                    if(!empty($content['offsetX'])) {
                        $drawing->setOffsetX($content['offsetX']);//设置X方向偏移量
                    }
                    if(!empty($content['offsetY'])) {
                        $drawing->setOffsetY($content['offsetY']);//设置Y方向偏移量
                    }

                    $drawing->setCoordinates($rowName.$this->_row);
                    $drawing->setWorksheet($this->workSheet);
                }
            }elseif(is_array($content) && isset($content['formula'])){
                // 新增：支持 ['formula' => '公式'] 写法
                $this->workSheet->setCellValueExplicit($rowName.$this->_row, $content['formula'], \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA);
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
     * @param $group_left_count
     * @author tinymeng
     */
    private function excelGroupLeft(array $data, $group_left_count)
    {
        // 获取分组字段在field中的实际位置
        $group_field_positions = [];
        foreach($this->group_left as $group_field){
            $position = array_search($group_field, $this->field);
            if($position !== false){
                $group_field_positions[] = $position;
            }
        }

        if(empty($group_field_positions)){
            throw new TinymengException(StatusCode::COMMON_PARAM_INVALID, '分组字段未在标题中定义');
        }

        $group_start = $this->_row;
        foreach ($data as $key => $val){
            // 第一级分组的合并单元格
            $rowName = $this->cellName($group_field_positions[0]); // 使用第一个分组字段的实际位置
            $coordinate = $rowName.$this->_row.':'.$rowName.($this->_row+$val['count']-1);
            $this->workSheet->mergeCells($coordinate);
            $this->workSheet->setCellValue($rowName.$this->_row, $key);

            // 新增：合并mergeColumns指定的其它列
            if (!empty($this->mergeColumns)) {
                foreach ($this->mergeColumns as $field) {
                    // 跳过分组字段本身
                    if (in_array($field, $this->group_left)) continue;
                    $colIdx = array_search($field, $this->field);
                    if ($colIdx !== false) {
                        $colLetter = $this->cellName($colIdx);
                        $this->workSheet->mergeCells($colLetter.$this->_row.':'.$colLetter.($this->_row+$val['count']-1));
                        // 取本组第一个数据的值
                        $this->workSheet->setCellValue($colLetter.$this->_row, $val['data'][0][$field] ?? '');
                    }
                }
            }
            
            if($group_left_count == 1){
                foreach ($val['data'] as $dataRow){
                    $this->excelSetCellValue($dataRow);
                }
            }else{
                $sub_group_start = $this->_row;
                $rowName = $this->cellName($group_field_positions[1]); // 使用第二个分组字段的实际位置
                
                foreach ($val['data'] as $k => $v){
                    $coordinate = $rowName.$sub_group_start.':'.$rowName.($sub_group_start+$v['count']-1);
                    $this->workSheet->mergeCells($coordinate);
                    $this->workSheet->setCellValue($rowName.$sub_group_start, $k);
                    
                    foreach ($v['data'] as $data){
                        $this->excelSetCellValue($data);
                    }
                    
                    $sub_group_start = $sub_group_start + $v['count'];
                }
            }
            
            $this->_row = $group_start + $val['count'];
            $group_start = $this->_row;
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
     * 自动合并指定字段相同值的单元格
     * @param int $rowStart 数据起始行
     * @param int $rowEnd 数据结束行
     */
    private function autoMergeColumns($rowStart, $rowEnd)
    {
        if ($rowEnd <= $rowStart) return;
        foreach ($this->mergeColumns as $fieldName) {
            $colIdx = array_search($fieldName, $this->field);
            if ($colIdx === false) continue;
            $colLetter = $this->cellName($colIdx);
            $lastValue = null;
            $mergeStart = $rowStart;
            for ($row = $rowStart; $row <= $rowEnd; $row++) {
                $cellValue = $this->workSheet->getCell($colLetter . $row)->getValue();
                if ($lastValue !== null && $cellValue !== $lastValue) {
                    if ($row - $mergeStart > 1) {
                        $this->workSheet->mergeCells($colLetter . $mergeStart . ':' . $colLetter . ($row - 1));
                    }
                    $mergeStart = $row;
                }
                $lastValue = $cellValue;
            }
            // 处理最后一组
            if ($rowEnd - $mergeStart + 1 > 1) {
                $this->workSheet->mergeCells($colLetter . $mergeStart . ':' . $colLetter . $rowEnd);
            }
        }
    }

    /**
     * 应用全表样式
     */
    private function applySheetStyle()
    {
        if (empty($this->sheetStyle)) return;
        // 计算数据区范围
        $startCol = 'A';
        $endCol = $this->cellName(count($this->field) - 1);
        $startRow = 1;
        $endRow = $this->_row - 1;
        $cellRange = $startCol . $startRow . ':' . $endCol . $endRow;
        $this->workSheet->getStyle($cellRange)->applyFromArray($this->sheetStyle);
    }

    /**
     * 设置自定义表格操作回调
     * @param callable $fn
     * @return $this
     */
    public function complexFormat(callable $fn) {
        $this->complexFormatCallback = $fn;
        return $this;
    }


}

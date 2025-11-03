<?php
/**
 * @name: TWorklSheet
 * @author: JiaMeng <666@majiameng.com>
 * @file: Export.php
 * @Date: 2024/03/04 10:15
 */
namespace tinymeng\spreadsheet\Excel;

use tinymeng\spreadsheet\Excel\Handler\CellValueHandler;
use tinymeng\spreadsheet\Excel\Handler\DataValidationHandler;
use tinymeng\spreadsheet\Excel\Handler\GroupHandler;
use tinymeng\spreadsheet\Excel\Handler\HeaderHandler;
use tinymeng\spreadsheet\Excel\Handler\MergeHandler;
use tinymeng\spreadsheet\Excel\Handler\StyleHandler;
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
     * 列的数据验证配置
     * @var array 格式：['field_name' => ['type' => 'list', 'options' => [...], 'promptTitle' => '', 'promptMessage' => '', ...]]
     */
    private $columnValidations = [];

    /**
     * 必填字段列表
     * @var array 格式：['field_name1', 'field_name2', ...] 或 在 titleConfig 中通过 'required_fields' 配置
     */
    private $requiredFields = [];

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
        $this->columnValidations = [];
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
            // 新增：读取必填字段配置
            if (isset($titleConfig['required_fields'])) {
                $this->requiredFields = $titleConfig['required_fields'];
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
            HeaderHandler::setHeader($this->workSheet, $this->mainTitle, $this->fileTitle);
            $this->_row++; // 当前行数
        }

        /** 设置表头 **/
        $result = HeaderHandler::setTitle(
            $this->workSheet,
            $this->fileTitle,
            $this->title_row ?? 1,
            $this->titleConfig,
            $this->_col,
            $this->_row,
            $this->titleHeight ?? null,
            $this->titleWidth ?? null,
            $this->requiredFields
        );
        $this->_col = $result['col'];
        $this->_row = $result['row'];

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
        // 读取样式配置
        if (!empty($this->config['subtotalStyle'])) {
            $this->subtotalStyle = $this->config['subtotalStyle'];
        }
        if (!empty($this->config['sheetStyle'])) {
            $this->sheetStyle = $this->config['sheetStyle'];
        }

        // 新增：应用全表样式
        StyleHandler::applySheetStyle(
            $this->workSheet,
            $this->sheetStyle,
            $this->field,
            $this->_row - 1
        );
        // 新增：应用数据验证（如果没有数据也要应用，用于导出模板）
        if (!empty($this->columnValidations) && empty($this->data)) {
            // 确定数据起始行
            $dataStartRow = $this->titleConfig['data_start_row'] ?? $this->_row;
            $this->applyAllColumnValidations($dataStartRow, 0);
        }
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
                MergeHandler::autoMergeColumns(
                    $this->workSheet,
                    $this->mergeColumns,
                    $this->field,
                    $rowStart,
                    $this->_row - 1
                );
            }
            // 新增：应用数据验证（无分组情况）
            if (!empty($this->columnValidations)) {
                $this->applyAllColumnValidations($rowStart, $this->_row - 1);
            }
        }else{   //根据设置分组字段进行分组
            /** 数据分组 **/
            $group_left_count = count($this->group_left);
            if($group_left_count == 1){
                $data = GroupHandler::groupDataByOneField($this->data, $this->group_left[0]);
                $rowStart = $this->_row;
                $this->_row = GroupHandler::processGroupLeft(
                    $this->workSheet,
                    $data,
                    $group_left_count,
                    $this->group_left,
                    $this->field,
                    $this->mergeColumns,
                    $this->_row,
                    function($val) {
                        return $this->excelSetCellValue($val);
                    }
                );
                // 新增：应用数据验证（分组情况1级）
                if (!empty($this->columnValidations)) {
                    $this->applyAllColumnValidations($rowStart, $this->_row - 1);
                }
            }elseif($group_left_count == 2){
                $this->data = GroupHandler::groupDataByTwoFields(
                    $this->data,
                    $this->group_left[0],
                    $this->group_left[1]
                );
                $rowStart = $this->_row;
                $this->_row = GroupHandler::processGroupLeft(
                    $this->workSheet,
                    $this->data,
                    $group_left_count,
                    $this->group_left,
                    $this->field,
                    $this->mergeColumns,
                    $this->_row,
                    function($val) {
                        return $this->excelSetCellValue($val);
                    }
                );
                // 新增：应用数据验证（分组情况2级）
                if (!empty($this->columnValidations)) {
                    $this->applyAllColumnValidations($rowStart, $this->_row - 1);
                }
            }else{
                throw new TinymengException(StatusCode::COMMON_PARAM_INVALID,
                    '左侧分组过多，导出失败！'
                );
            }
        }
    }


    /**
     * excel单元格赋值
     * @author tinymeng
     * @param array $val 数据
     */
    private function excelSetCellValue($val)
    {
        $this->_row = CellValueHandler::setCellValue(
            $this->workSheet,
            $val,
            $this->field,
            $this->_row,
            $this->title_row ?? 1,
            $this->height ?? null,
            $this->autoDataType ?? false,
            $this->format ?? true,
            $this->format_date ?? 'Y-m-d H:i:s'
        );
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

    /**
     * 设置列的数据验证和输入提示
     * @param string $fieldName 字段名（对应 $fileTitle 中的字段）
     * @param array $config 验证配置
     *   格式：
     *   [
     *     'type' => 'list',           // 验证类型: list(下拉列表), whole(整数), decimal(小数), date(日期), time(时间), textLength(文本长度), custom(自定义公式)
     *     'options' => ['选项1', '选项2'], // 当type为list时的选项列表
     *     'formula' => 'A1:A10',      // 当type为custom或list需要引用范围时的公式
     *     'operator' => 'between',    // 操作符: between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual
     *     'min' => 0,                // 最小值（用于whole, decimal, date, time, textLength）
     *     'max' => 100,              // 最大值（用于whole, decimal, date, time, textLength）
     *     'promptTitle' => '输入提示',   // 输入提示标题
     *     'promptMessage' => '请输入...', // 输入提示内容
     *     'errorTitle' => '输入错误',    // 错误提示标题
     *     'errorMessage' => '输入值无效',  // 错误提示内容
     *     'errorStyle' => 'stop',     // 错误样式: stop(停止), warning(警告), information(信息)
     *     'showInputMessage' => true, // 是否显示输入提示
     *     'showErrorMessage' => true,  // 是否显示错误提示
     *     'allowBlank' => false,       // 是否允许空白
     *     'showDropDown' => true,      // 是否显示下拉箭头（仅list类型）
     *     'data_end_row' => 200,       // 数据结束行（模板导出时使用，为0或未设置时应用到整列）
     *     'data_row_count' => 100,     // 数据行数（模板导出时使用，从数据起始行开始计算的行数）
     *   ]
     * @return $this
     */
    public function setColumnValidation(string $fieldName, array $config) {
        $this->columnValidations[$fieldName] = $config;
        return $this;
    }

    /**
     * 批量设置列的数据验证
     * @param array $validations 格式：['field_name' => [验证配置], ...]
     * @return $this
     */
    public function setColumnValidations(array $validations) {
        foreach ($validations as $fieldName => $config) {
            $this->setColumnValidation($fieldName, $config);
        }
        return $this;
    }

    /**
     * 应用数据验证到指定列
     * @param string $fieldName 字段名
     * @param int $startRow 起始行（数据开始行）
     * @param int $endRow 结束行（数据结束行，为0时表示应用到整列）
     */
    private function applyColumnValidation(string $fieldName, int $startRow, int $endRow = 0) {
        if (!isset($this->columnValidations[$fieldName])) {
            return;
        }

        // 查找字段在 field 数组中的位置
        $fieldIndex = array_search($fieldName, $this->field);
        if ($fieldIndex === false) {
            return;
        }

        $config = $this->columnValidations[$fieldName];

        // 使用 DataValidationHandler 处理验证
        DataValidationHandler::applyValidation(
            $this->workSheet,
            $fieldName,
            $config,
            $fieldIndex,
            $startRow,
            $endRow,
            empty($this->data)
        );
    }

    /**
     * 应用所有列的数据验证
     * @param int $dataStartRow 数据开始行
     * @param int $dataEndRow 数据结束行（为0时表示应用到整列）
     */
    private function applyAllColumnValidations(int $dataStartRow, int $dataEndRow = 0) {
        foreach ($this->columnValidations as $fieldName => $config) {
            $this->applyColumnValidation($fieldName, $dataStartRow, $dataEndRow);
        }
    }

    /**
     * 设置必填字段
     * @param array $fields 必填字段列表，格式：['field_name1', 'field_name2', ...]
     * @return $this
     */
    public function setRequiredFields(array $fields) {
        $this->requiredFields = $fields;
        return $this;
    }

    /**
     * 获取必填字段列表
     * @return array
     */
    public function getRequiredFields(): array {
        return $this->requiredFields;
    }

}

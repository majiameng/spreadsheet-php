<?php
use tinymeng\spreadsheet\TSpreadSheet;
use tinymeng\spreadsheet\Util\ConstCode;

require __DIR__.'/../vendor/autoload.php';

/**
 * excel生成文件名
 */
$filename = $sheetName = "export_demo";
/**
 * excel表头
 */
$title = [
    '序号'=>'_id',
    'ID'=>'goods_id',
    '状态'=>'status',
];


// excel字段验证
$columnValidation = [
    'goods_id'=>[
        'type' => 'whole',
        'min' => 1,
        'max' => 10000000,
        'promptMessage' => '请输入ID',
        'errorMessage' => '只能输入数字',
    ],
    'status'=>[
        'type' => 'list',
        'options' => ['开启','关闭'],
        'promptTitle' => '请选择',
        'promptMessage' => '请从下拉列表中选择',
        'errorMessage' => '只能从下拉列表中选择',
        'showDropDown' => true
    ],
];
// 配置参数
$config = [
    'horizontalCenter' => true,               // 是否居中
    'titleHeight' => 22,                      // 定义表头行高
    'titleWidth' => 20,                       // 定义表头列宽
    'height' => 22,                           // 定义数据行高
    'autoFilter' => true,                     // 开启自动筛选
    'freezePane' => false,                    // 冻结窗格（首行首列）
    'fieldMappingMethod' => ConstCode::FIELD_MAPPING_METHOD_NAME_CORRESPONDING_FIELD,  // 名称对应字段方式
];

$TSpreadSheet = TSpreadSheet::export($config)
    //创建一个sheet，设置sheet表头，并给表格赋值
    ->createWorkSheet($sheetName);
foreach ($columnValidation as $key => $value){
    $TSpreadSheet->setColumnValidation($key,$value);
}
//文件存储本地
$path = $TSpreadSheet->setWorkSheetData($title,[])->generate()->save($filename);
echo '生成excel路径：'.$path;
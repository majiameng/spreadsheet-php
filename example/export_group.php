<?php
use tinymeng\spreadsheet\TSpreadSheet;
use tinymeng\spreadsheet\Util\ConstCode;

require __DIR__.'/../vendor/autoload.php';

/**
 * excel生成文件名
 */
$filename = $sheetName = "export_group_demo";
/**
 * excel表头
 * 注意：分组字段必须在title中定义
 */
$titleConfig = [
    'title_row' => 2,  // 表头行号
    'group_left' => ['user_id', 'day'],  // 左侧分组字段，最多支持两级分组
    'title' => [
        'ID' => 'id',
        '用户ID' => 'user_id',    // 分组字段1
        '结算日期' => 'day',      // 分组字段2
        '订单编号' => 'order_sn',
        '下单时间' => 'create_time',
    ]
];

/**
 * excel数据数组（二维）
 */
$data = [
    [
        'id' => '1',
        'user_id' => '1000',      // 第一组用户
        'day' => '20220101',      // 第一天
        'order_sn' => '20180101465464',
        'create_time' => '1687140376',
    ],
    [
        'id' => '2',
        'user_id' => '1000',      // 第一组用户
        'day' => '20220101',      // 第一天
        'order_sn' => '20180101465465',
        'create_time' => '1687140377',
    ],
    [
        'id' => '3',
        'user_id' => '1000',      // 第一组用户
        'day' => '20220102',      // 第二天
        'order_sn' => '20180102465466',
        'create_time' => '1687140378',
    ],
    [
        'id' => '4',
        'user_id' => '1001',      // 第二组用户
        'day' => '20220101',      // 第一天
        'order_sn' => '20180101465467',
        'create_time' => '1687140379',
    ],
    [
        'id' => '5',
        'user_id' => '1000',      // 第二组用户
        'day' => '20220101',      // 第一天
        'order_sn' => '20180101465468',
        'create_time' => '1687140379',
    ],
    [
        'id' => '6',
        'user_id' => '1001',      // 第二组用户
        'day' => '20220101',      // 第一天
        'order_sn' => '20180101465469',
        'create_time' => '1687140379',
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

// 创建导出实例并设置数据
$TSpreadSheet = TSpreadSheet::export($config)
    ->createWorkSheet($sheetName)
    ->setWorkSheetData($titleConfig, $data);

// 生成并保存文件
$path = $TSpreadSheet->generate()->save($filename);
echo '生成excel路径：'.$path;
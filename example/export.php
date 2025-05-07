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
//config 中 'fieldMappingMethod'=>ConstCode::FIELD_MAPPING_METHOD_NAME_CORRESPONDING_FIELD,//名称对应字段
$title = [
    '序号'=>'_id',
    'ID'=>'id',
    '订单编号'=>'order_sn',
    '用户id'=>'user_id',
    '结算日期'=>'day',
    '下单时间'=>'create_time',
    '图片'=>'image',
];

//config 中 'fieldMappingMethod'=>ConstCode::FIELD_MAPPING_METHOD_FIELD_CORRESPONDING_NAME,//字段对应名称
//$title = [
//    '_id'=>'序号',
//    'id'=>'ID',
//    'order_sn'=>'订单编号',
//    'user_id'=>'用户id',
//    'day'=>'结算日期',
//    'create_time'=>'下单时间',
//    'image'=>'图片',
//];

/**
 * excel数据数组（二维）
 */
$data = [
    [
        'id'=>'1',
        'order_sn'=>'20180101465464',
        'user_id'=>'1000',
        'day'=>'20220101',
        'create_time'=>'1687140376',
        'image'=>[
            'type'=>'image',
            'content'=>'http://static.majiameng.com/main/img/portrait.jpg',//网络图片确保存在
            'height'=>100,
//            'width'=>100,//只设置高，宽会自适应，如果设置宽后，高则失效
        ],
    ],[
        'id'=>'2',
        'order_sn'=>'20190101465464',
        'user_id'=>'1000',
        'day'=>'20220101',
        'create_time'=>'1687140376',
        'image'=>[
            'type'=>'image',
            'content'=>'./text.png',//本地图片确保存在
            'height'=>100,
        ],
    ],[
        'id'=>'3',
        'order_sn'=>'20200101465464',
        'user_id'=>'1000',
        'day'=>'20220101',
        'create_time'=>'1687140376',
    ],[
        'id'=>'4',
        'order_sn'=>'20210101465464',
        'user_id'=>'1001',
        'day'=>'20220101',
        'create_time'=>'1687140376',
    ],
];
$TSpreadSheet = TSpreadSheet::export()
    //创建一个sheet，设置sheet表头，并给表格赋值
    ->createWorkSheet($sheetName)->setWorkSheetData($title,$data);
//    ->createWorkSheet($sheetName1)->setWorkSheetData($title1,$data1);//如果多个sheet可多次创建

//文件存储本地
$path = $TSpreadSheet->generate()->save($filename);
echo '生成excel路径：'.$path;exit();
//生成excel路径：E:\spreadsheet-php\example\public\export\20240402\export_demo_2024-04-02_351.xlsx

//这样直接输出到浏览器中下载
$TSpreadSheet->generate()->download($filename);

//配置参数可以通过配置文件在初始化时传入
$config = [
    'pathName'=>null,                       //文件存储位置
    'fileName'=>null,                       //文件名称
    'horizontalCenter'=>true,               //是否居中
    'titleHeight'=>null,                    //定义表头行高,常用22
    'titleWidth'=>null,                     //定义表头列宽(未设置则自动计算宽度),常用20
    'height'=>null,                         //定义数据行高,常用22
    'autoFilter'=>false,                    //自动筛选(是否开启)
    'autoDataType'=>true,                   //自动适应文本类型
    'freezePane'=>false,                    //冻结窗格（要冻结的首行首列"B2"，false不开启）
    'fieldMappingMethod'=>ConstCode::FIELD_MAPPING_METHOD_NAME_CORRESPONDING_FIELD,//字段映射方式
];
$TSpreadSheet = TSpreadSheet::export($config);
//配置参数也可以后期赋值
$TSpreadSheet = TSpreadSheet::export($config)->setAutoFilter(true);
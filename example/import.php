<?php
use tinymeng\spreadsheet\TSpreadSheet;

require __DIR__.'/vendor/autoload.php';

/**
 * excel生成文件名
 */
$filename = './export_demo.xlsx';
/**
 * excel表头
 */
$title = [
    '序号'=>'id',
    '订单编号'=>'order_sn',
    '用户id'=>'user_id',
    '结算日期'=>'day',
    '下单时间'=>'create_time',
];

//读取并初始化表格内容数据
$TSpreadSheet = TSpreadSheet::import()
    ->setFileName($filename)//读取文件路径
    ->initWorkSheet($filename);//读取并初始化表格内容数据

//设置title对应字段,获取表格内容
$data = $TSpreadSheet->setTitle($title)->getExcelData();
var_dump($data);die;
/**
 * array(3) {
 * [0]=>
 * array(3) {
 * ["id"]=>
 * string(1) "1"
 * ["order_sn"]=>
 * string(14) "20180101465464"
 * ["create_time"]=>
 * string(19) "2023-06-19 10:06:16"
 * }
 * [1]=>
 * array(3) {
 * ["id"]=>
 * string(1) "2"
 * ["order_sn"]=>
 * string(14) "20190101465464"
 * ["create_time"]=>
 * string(19) "2023-06-19 10:06:16"
 * }
 * [2]=>
 * array(3) {
 * ["id"]=>
 * string(1) "3"
 * ["order_sn"]=>
 * string(14) "20200101465464"
 * ["create_time"]=>
 * string(19) "2023-06-19 10:06:16"
 * }
 * }
 */


//也可以设置读取第几个sheet
$TSpreadSheet = TSpreadSheet::import()
    ->setFileName($filename)
    ->setSheet(0)//读取第0个sheet
    ->setTitleRow(1)//表头所在行
    ->initWorkSheet($filename);
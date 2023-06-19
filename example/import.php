<?php
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

/**
 * 初始化导入类
 * @var  $importUtil
 */
$export = new \tinymeng\spreadsheet\Import();
$export->fileName = $filename;
$export->sheet = 0;//第一个sheet
$export->titleFieldsRow = 1;//表头所在行

/** 读取并初始化表格内容数据 */
$export->initReadExcel();

/** 设置title对应字段 */
$export->setTitle($title);
/** 获取表头字段 */
$fields = $export->getTitleFields();

/** 获取表格内容 */
$data = $export->getExcelData($fields);
var_dump($data);

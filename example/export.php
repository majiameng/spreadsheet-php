<?php
require __DIR__.'/vendor/autoload.php';

use tinymeng\spreadsheet\Export;

/**
 * excel生成文件名
 */
$filename = "export_demo";
/**
 * excel表头
 */
$title = [
    '序号'=>'id',
    '订单编号'=>'order_sn',
    '用户id'=>'user_id',
    '结算日期'=>'day',
    '下单时间'=>'create_time',
    '图片'=>'image',
];

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
            'content'=>'https://sns.bjwmsc.com/wp-content/themes/zibll/img/logo.png',//网络图片确保存在
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

$fileTitle = [
    'title_row' => 1,
    'title' => $title,
];
$export = new Export();
$export->fileTitle = $fileTitle;//表头
$export->sheetName = $filename;//文件名
$export->data = $data;//excel数据数组
$export->freezePane = false;
$export->height = 70;//默认是22，如果有图片适当调高些
$export->saveType = 'save';//存储方式: download下载, save存储本地
$path = $export->exportExcel();
echo '生成excel路径：'.$path;

<h1 align="center">tinymeng/spreadsheet</h1>

欢迎 Star，欢迎 PR！

> 大家如果有问题要交流，就发在这里吧： [Spreadsheet](https://github.com/majiameng/spreadsheet-php/issues/1) 交流 或发邮件 666@majiameng.com


# PHP Spreadsheet Class


## 1.安装
> composer require tinymeng/spreadsheet  -vvv


* 2.1 excel导出
* 2.2 excel导入

#### 2.1.excel导出

> example 事例目录


```php
<?php

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
    '下单时间'=>'create_time',
];

/**
 * excel数据数组（二维）
 */
$data = [
    [
        'id'=>'1',
        'order_sn'=>'20180101465464',
        'create_time'=>'1687140376',
    ],[
        'id'=>'2',
        'order_sn'=>'20190101465464',
        'create_time'=>'1687140376',
    ],[
        'id'=>'3',
        'order_sn'=>'20200101465464',
        'create_time'=>'1687140376',
    ],
];

$fileTitle = [
    'title_row' => 1,
    'title' => $title,
];
$export = new \tinymeng\spreadsheet\Export();
$export->fileTitle = $fileTitle;//表头
$export->sheetName = $filename;//文件名
$export->data = $data;//excel数据数组
$export->saveType = 'save';//存储方式: download下载, save存储本地
$path = $export->exportExcel();
echo '生成excel路径：'.$path;
//生成excel路径：E:\public\export\20230619\export_demo_2023-06-19_567.xlsx
```

#### 2.2.excel导入

```php
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

```
打印结果如下
```
array(3) {
  [0]=>
  array(3) {
    ["id"]=>
    string(1) "1"
    ["order_sn"]=>
    string(14) "20180101465464"
    ["create_time"]=>
    string(19) "2023-06-19 10:06:16"
  }
  [1]=>
  array(3) {
    ["id"]=>
    string(1) "2"
    ["order_sn"]=>
    string(14) "20190101465464"
    ["create_time"]=>
    string(19) "2023-06-19 10:06:16"
  }
  [2]=>
  array(3) {
    ["id"]=>
    string(1) "3"
    ["order_sn"]=>
    string(14) "20200101465464"
    ["create_time"]=>
    string(19) "2023-06-19 10:06:16"
  }
}
```

> 大家如果有问题要交流，就发在这里吧： [Spreadsheet](https://github.com/majiameng/spreadsheet-php/issues/1) 交流 或发邮件 666@majiameng.com

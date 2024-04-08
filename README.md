<h1 align="center">tinymeng/spreadsheet</h1>

欢迎 Star，欢迎 PR！

> 大家如果有问题要交流，就发在这里吧： [Spreadsheet](https://github.com/majiameng/spreadsheet-php/issues/1) 交流 或发邮件 666@majiameng.com


# PHP Spreadsheet Class
基于 `phpoffice/phpspreadsheet` 扩展快速封装使用，避免我们再重复性的造轮子

## 您可以在网站上找到tinymeng/spreadsheet文档。查看“入门”页面以获取快速概述。

* [Wiki Home](https://github.com/majiameng/spreadsheet-php/wiki)
* [中文文档](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Home)
* [开始](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Getting-Started)
* [安装](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Installation)
* [配置文件](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Configuration)
* [贡献指南](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Contributing-Guide)
* [更新日志](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Update-log)

# Installation

```
composer require tinymeng/spreadsheet  -vvv
```

> 类库使用的命名空间为 `\\tinymeng\\spreadsheet`

* [开始入门](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Getting-Started)

### 目录结构

```
.
├── config                          配置文件目录
│   └── TSpreadSheet.php            
├── example                         事例代码
│   ├── export.php                  导出事例代码
│   └── import.php                  导入事例代码
├── src                             代码源文件目录
│   ├── Connector
│   │   ├── Gateway.php             必须继承的抽象类
│   │   └── GatewayInterface.php    必须实现的接口
│   ├── Gateways
│   │   ├── Export.php              导出实例
│   │   └── Import.php              导入实例
│   ├── Util
│   │   └── TConfig.php             配置类
│   └── TSpreadSheet.php            抽象实例类
├── composer.json                   Composer File
├── LICENSE                         MIT License
├── README_zh_cn.md                 中文文档
└── README.md                       Documentation
```


### Configuration
[Configuration](https://github.com/majiameng/spreadsheet-php/wiki/zh-cn-Configuration)


* 2.1 excel导出 TSpreadSheet::export()
* 2.2 excel导入 TSpreadSheet::import()

#### 2.1.excel导出

> example 事例目录


```php
<?php
use tinymeng\spreadsheet\TSpreadSheet;

require __DIR__.'/vendor/autoload.php';

/**
 * excel生成文件名
 */
$filename = $sheetName = "export_demo";
/**
 * excel表头
 */
$title = [
    '序号'=>'_id',
    'ID'=>'id',
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
$TSpreadSheet = TSpreadSheet::export()
    //创建一个sheet，设置sheet表头，并给表格赋值
    ->createWorkSheet($sheetName)->setWorkSheetData($title,$data);
//    ->createWorkSheet($sheetName1)->setWorkSheetData($title1,$data1);//如果多个sheet可多次创建

//文件存储本地
$path = $TSpreadSheet->generate()->save($filename);
echo '生成excel路径：'.$path;exit();
//生成excel路径：E:\spreadsheet-php\example\public\export\20240402\export_demo_2024-04-02_351.xlsx
```
这样直接输出到浏览器中下载
```
$TSpreadSheet->generate()->download($filename);
```
配置参数可以通过配置文件在初始化时传入
```
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
];
$TSpreadSheet = TSpreadSheet::export($config);
//配置参数也可以后期赋值
$TSpreadSheet = TSpreadSheet::export($config)->setAutoFilter(true);
```

#### 2.2.excel导入

```php
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
    ->initWorkSheet($filename);//读取并初始化表格内容数据

//设置title对应字段,获取表格内容
$data = $TSpreadSheet->setTitle($title)->getExcelData();
var_dump($data);die;

//也可以设置读取第几个sheet
$TSpreadSheet = TSpreadSheet::import()
    ->setFileName($filename)//读取文件路径
    ->setSheet(0)//读取第0个sheet
    ->setTitleRow(1)//表头所在行
    ->initWorkSheet();

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

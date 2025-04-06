<?php
use \tinymeng\spreadsheet\Util\ConstCode;
return [
    'creator'=>'tinymeng',                  //文件创建者
    'pathName'=>null,                       //文件存储位置
    'fileName'=>null,                       //文件名称
    'horizontalCenter'=>true,               //是否居中
    'titleHeight'=>null,                    //定义表头行高,常用22
    'titleWidth'=>null,                     //定义表头列宽(未设置则自动计算宽度),常用20
    'height'=>null,                         //定义数据行高,常用22
    'autoFilter'=>false,                    //自动筛选(是否开启)
    'autoDataType'=>true,                   //自动适应文本类型
    'freezePane'=>false,                    //冻结窗格（要冻结的首行首列"B2"，false不开启）
    'fieldMappingMethod'=>ConstCode::FIELD_MAPPING_METHOD_FIELD_CORRESPONDING_NAME,//字段映射方式
];

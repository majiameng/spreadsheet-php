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
    /**
     * 字段映射方式
     * ConstCode::FIELD_MAPPING_METHOD_FIELD_CORRESPONDING_NAME = 1;//字段对应名称
     * ConstCode::FIELD_MAPPING_METHOD_NAME_CORRESPONDING_FIELD = 2;//名称对应字段
     */
    'fieldMappingMethod'=>ConstCode::FIELD_MAPPING_METHOD_NAME_CORRESPONDING_FIELD,
    'mainTitleLine'=>false,                 //主标题行是否显示
    'mainTitle'=>'',                        //主标题名称，默认为sheet的名称
];

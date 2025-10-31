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
    'title_row' => 2,  // 表头占用行数
    'title_start_row' => null,  // 表头开始行数
    'group_left' => ['id'], // 以id分组
    'mergeColumns' => ['meeting_name', 'time'], // 需要自动合并的字段
    'title' => [
        '场次编号' => 'id',
        '会议名称' => 'meeting_name',
        '会议时间' => 'time',
        '姓名' => 'turename',
        '服务费' => 'price',
    ]
];

/**
 * excel数据数组（二维）
 */
$data = [
    [
        'id' => '1',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张三',
        'price' => '500',
    ],
    [
        'id' => '1',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张1',
        'price' => '5040',
    ],
    [
        'id' => '1',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张2',
        'price' => '53300',
    ],
    [
        'id' => '2',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张333',
        'price' => '53300',
    ],
    [
        'id' => '2',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张4444',
        'price' => '53300',
    ],
    [
        'id' => '2',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张4444',
        'price' => '53300',
    ],
    [
        'id' => '3',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张4444',
        'price' => '53300',
    ],
    [
        'id' => '4',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张4444',
        'price' => '53300',
    ],
    [
        'id' => '5',
        'meeting_name' => '会议名称',
        'time' => '2025年4月2日',
        'turename' => '张4444',
        'price' => '53300',
    ],
];

// 配置参数
$config = [
    'horizontalCenter' => true,               // 是否居中
    'titleHeight' => 22,                      // 定义表头行高
    'titleWidth' => 20,                       // 定义表头列宽
    'height' => 22,                           // 定义数据行高
    'autoFilter' => false,                     // 开启自动筛选
    'freezePane' => false,                    // 冻结窗格（首行首列）
    'fieldMappingMethod' => ConstCode::FIELD_MAPPING_METHOD_NAME_CORRESPONDING_FIELD,  // 名称对应字段方式
    // 更多样式请查询 https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#styling-cells
    'sheetStyle' => [
        'font' => [
            'name' => '微软雅黑',
            'size' => 8,
        ],
        'alignment' => [
            'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
        ],
    ],
];

$complexFormat = function($sheet) {

    // 第1行：大标题
    $sheet->mergeCells('A1:E1');
    $sheet->setCellValue('A1', '非常之路-慢性病诊疗方案公开课-会议执行总结报告');
    $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(16);
    $sheet->getStyle('A1')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

    // 第2行：导出时间
    $sheet->mergeCells('A13:E13');
    $sheet->setCellValue('A13', '导出时间：' . date('Y-m-d'));
    $sheet->getStyle('A13')->getFont()->setSize(10)->getColor()->setRGB('888888');
    $sheet->getStyle('A13')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT);

    // 设置A1字体加粗、字号14、红色（演示，实际A1已设置16号字）
    $sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14)->getColor()->setRGB('FF0000');
    // 设置A2:E2背景色
    $sheet->getStyle('A13:E13')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setRGB('FFFFCC');
    // 你可以继续添加更多自定义操作
};

// 创建导出实例并设置数据
$TSpreadSheet = TSpreadSheet::export($config)
    ->createWorkSheet($sheetName)
    ->setMainTitle("北京222公司") //设置大标题
    ->complexFormat($complexFormat)
    ->setWorkSheetData($titleConfig, $data);

// 生成并保存文件
$path = $TSpreadSheet->generate()->save($filename);
echo '生成excel路径：'.$path;
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
$startRow = 4;
$titleConfig = [
    'title_show' => true,
    'data_start_row' => $startRow,  // 内容开始行数
    'group_left' => ['id'], // 以id分组
    'mergeColumns' => ['meeting_name', 'time'], // 需要自动合并的字段
    'title' => [
        'id', 'meeting_name', 'time', 'turename', 'price',
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

$data[] = [
    'id' => '合计',
    'price' => ['formula' => '=SUM(E'.$startRow.':E'.($startRow+count($data)-1).')'],
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

$complexFormat = function($sheet) use ($startRow,$data) {
    $endRow = $startRow + count($data) - 1;
    $cellRange = "A{$startRow}:E{$endRow}";

    // 设置边框
    $sheet->getStyle($cellRange)->getBorders()->getAllBorders()->setBorderStyle(
        \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN
    );
};



// 模板文件（带有已设置好格式的文件）
$filePath = './test.xlsx';

// 创建导出实例并设置数据
$TSpreadSheet = TSpreadSheet::export($config)
    ->setMainTitle("北京222公司") //设置大标题
    ->loadFile($filePath)// 加载模板文件
    ->selectWorkSheet()// 选择工作表
    ->complexFormat($complexFormat)
    ->setWorkSheetData($titleConfig, $data);

// 生成并保存文件
$path = $TSpreadSheet->generate()->save($filename);
echo '生成excel路径：'.$path;
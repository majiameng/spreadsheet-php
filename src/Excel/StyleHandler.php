<?php
/**
 * @name: StyleHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: StyleHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class StyleHandler
{
    /**
     * 应用全表样式
     * @param Worksheet $worksheet
     * @param array $sheetStyle 样式配置数组
     * @param array $fields 字段列表
     * @param int $endRow 数据结束行
     * @param callable $cellNameFunc 获取列字母的函数
     */
    public static function applySheetStyle(
        Worksheet $worksheet,
        array $sheetStyle,
        array $fields,
        int $endRow,
        callable $cellNameFunc
    ) {
        if (empty($sheetStyle)) return;
        
        // 计算数据区范围
        $startCol = 'A';
        $endCol = $cellNameFunc(count($fields) - 1);
        $startRow = 1;
        $cellRange = $startCol . $startRow . ':' . $endCol . $endRow;
        $worksheet->getStyle($cellRange)->applyFromArray($sheetStyle);
    }
}


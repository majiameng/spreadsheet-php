<?php
/**
 * @name: StyleHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: StyleHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel\Handler;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use tinymeng\spreadsheet\Util\WorkSheetHelper;

class StyleHandler
{
    /**
     * 应用全表样式
     * @param Worksheet $worksheet
     * @param array $sheetStyle 样式配置数组
     * @param array $fields 字段列表
     * @param int $endRow 数据结束行
     */
    public static function applySheetStyle(
        Worksheet $worksheet,
        array $sheetStyle,
        array $fields,
        int $endRow
    ) {
        if (empty($sheetStyle)) return;
        
        // 计算数据区范围
        $startCol = 'A';
        $endCol = WorkSheetHelper::cellName(count($fields) - 1);
        $startRow = 1;
        $cellRange = $startCol . $startRow . ':' . $endCol . $endRow;
        $worksheet->getStyle($cellRange)->applyFromArray($sheetStyle);
    }
}


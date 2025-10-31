<?php
/**
 * @name: MergeHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: MergeHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel\Handler;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use tinymeng\spreadsheet\Util\WorkSheetHelper;

class MergeHandler
{
    /**
     * 自动合并指定字段相同值的单元格
     * @param Worksheet $worksheet
     * @param array $mergeColumns 需要合并的列字段名列表
     * @param array $fields 所有字段列表
     * @param int $rowStart 数据起始行
     * @param int $rowEnd 数据结束行
     */
    public static function autoMergeColumns(
        Worksheet $worksheet,
        array $mergeColumns,
        array $fields,
        int $rowStart,
        int $rowEnd
    ) {
        if ($rowEnd <= $rowStart) return;
        
        foreach ($mergeColumns as $fieldName) {
            $colIdx = array_search($fieldName, $fields);
            if ($colIdx === false) continue;
            
            $colLetter = WorkSheetHelper::cellName($colIdx);
            $lastValue = null;
            $mergeStart = $rowStart;
            
            for ($row = $rowStart; $row <= $rowEnd; $row++) {
                $cellValue = $worksheet->getCell($colLetter . $row)->getValue();
                if ($lastValue !== null && $cellValue !== $lastValue) {
                    if ($row - $mergeStart > 1) {
                        $worksheet->mergeCells($colLetter . $mergeStart . ':' . $colLetter . ($row - 1));
                    }
                    $mergeStart = $row;
                }
                $lastValue = $cellValue;
            }
            
            // 处理最后一组
            if ($rowEnd - $mergeStart + 1 > 1) {
                $worksheet->mergeCells($colLetter . $mergeStart . ':' . $colLetter . $rowEnd);
            }
        }
    }
}


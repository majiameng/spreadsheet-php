<?php
/**
 * @name: HeaderHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: HeaderHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class HeaderHandler
{
    /**
     * 设置主标题（第一行）
     * @param Worksheet $worksheet
     * @param string $mainTitle
     * @param array $fileTitle
     * @param callable $cellNameFunc
     * @return int 返回标题所在行号
     */
    public static function setHeader(Worksheet $worksheet, string $mainTitle, array $fileTitle, callable $cellNameFunc): int
    {
        $row = 1;
        if (!empty($mainTitle)) {
            $worksheet->setCellValue('A' . $row, $mainTitle);
            
            // 计算实际的标题列数
            $titleCount = 0;
            foreach ($fileTitle as $val) {
                if (is_array($val)) {
                    $titleCount += count($val); // 如果是数组，加上子项的数量
                } else {
                    $titleCount++; // 如果是单个标题，加1
                }
            }
            
            // 使用实际的标题列数来合并单元格
            $worksheet->mergeCells('A' . $row . ':' . $cellNameFunc($titleCount - 1) . $row);
        }
        return $row;
    }

    /**
     * 设置表头
     * @param Worksheet $worksheet
     * @param array $fileTitle
     * @param int $titleRow 标题行数（合并行数）
     * @param array $titleConfig 标题配置
     * @param int $col 当前列索引
     * @param int $row 当前行索引
     * @param callable $cellNameFunc
     * @param int|null $titleHeight 行高
     * @param int|null $titleWidth 列宽
     * @return array ['col' => int, 'row' => int] 返回更新后的列和行
     */
    public static function setTitle(
        Worksheet $worksheet,
        array $fileTitle,
        int $titleRow,
        array $titleConfig,
        int $col,
        int $row,
        callable $cellNameFunc,
        ?int $titleHeight = null,
        ?int $titleWidth = null
    ): array {
        if (!empty($titleConfig['title_start_row'])) {
            $row = $titleConfig['title_start_row'];
        }

        $_merge = $cellNameFunc($col);
        foreach ($fileTitle as $key => $val) {
            if (!empty($titleHeight)) {
                $worksheet->getRowDimension($col)->setRowHeight($titleHeight); // 行高度
            }
            $rowName = $cellNameFunc($col);
            $worksheet->getStyle($rowName . $row)->getAlignment()->setWrapText(true); // 自动换行
            
            if (is_array($val)) {
                $num = 1;
                $_cols = $col;
                foreach ($val as $k => $v) {
                    if (!isset($titleConfig['title_show']) || $titleConfig['title_show'] !== false) {
                        $worksheet->setCellValue($cellNameFunc($_cols) . ($row + 1), $k);
                    }
                    if (!empty($titleWidth)) {
                        $worksheet->getColumnDimension($cellNameFunc($_cols))->setWidth($titleWidth); // 列宽度
                    } else {
                        $worksheet->getColumnDimension($cellNameFunc($_cols))->setAutoSize(true); // 自动计算宽度
                    }
                    if ($num < count($val)) {
                        $col++;
                        $num++;
                    }
                    $_cols++;
                }
                $worksheet->mergeCells($_merge . $row . ':' . $cellNameFunc($col) . $row);
                if (!isset($titleConfig['title_show']) || $titleConfig['title_show'] !== false) {
                    $worksheet->setCellValue($_merge . $row, $key); // 设置值
                }
            } else {
                if ($titleRow != 1) {
                    $worksheet->mergeCells($rowName . $row . ':' . $rowName . ($row + $titleRow - 1));
                }
                if (!isset($titleConfig['title_show']) || $titleConfig['title_show'] !== false) {
                    $worksheet->setCellValue($rowName . $row, $key); // 设置值
                }
                if (!empty($titleWidth)) {
                    $worksheet->getColumnDimension($rowName)->setWidth($titleWidth); // 列宽度
                } else {
                    $worksheet->getColumnDimension($rowName)->setAutoSize(true); // 自动计算宽度
                }
            }
            $col++;
            $_merge = $cellNameFunc($col);
        }
        $row += $titleRow; // 当前行数
        
        return ['col' => $col, 'row' => $row];
    }
}


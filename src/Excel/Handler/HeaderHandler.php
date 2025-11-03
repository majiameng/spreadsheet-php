<?php
/**
 * @name: HeaderHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: HeaderHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel\Handler;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Style\Color;
use tinymeng\spreadsheet\Util\WorkSheetHelper;

class HeaderHandler
{
    /**
     * 设置主标题（第一行）
     * @param Worksheet $worksheet
     * @param string $mainTitle
     * @param array $fileTitle
     * @return int 返回标题所在行号
     */
    public static function setHeader(Worksheet $worksheet, string $mainTitle, array $fileTitle): int
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
            $worksheet->mergeCells('A' . $row . ':' . WorkSheetHelper::cellName($titleCount - 1) . $row);
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
     * @param int|null $titleHeight 行高
     * @param int|null $titleWidth 列宽
     * @param array $requiredFields 必填字段列表
     * @return array ['col' => int, 'row' => int] 返回更新后的列和行
     */
    public static function setTitle(
        Worksheet $worksheet,
        array $fileTitle,
        int $titleRow,
        array $titleConfig,
        int $col,
        int $row,
        ?int $titleHeight = null,
        ?int $titleWidth = null,
        array $requiredFields = []
    ): array {
        if (!empty($titleConfig['title_start_row'])) {
            $row = $titleConfig['title_start_row'];
        }

        $_merge = WorkSheetHelper::cellName($col);
        foreach ($fileTitle as $key => $val) {
            if (!empty($titleHeight)) {
                $worksheet->getRowDimension($col)->setRowHeight($titleHeight); // 行高度
            }
            $rowName = WorkSheetHelper::cellName($col);
            $worksheet->getStyle($rowName . $row)->getAlignment()->setWrapText(true); // 自动换行
            
            if (is_array($val)) {
                $num = 1;
                $_cols = $col;
                foreach ($val as $k => $v) {
                    if (!isset($titleConfig['title_show']) || $titleConfig['title_show'] !== false) {
                        // 检查子标题是否必填
                        $cellValue = self::formatHeaderTitle($k, $v, $requiredFields);
                        $worksheet->setCellValue(WorkSheetHelper::cellName($_cols) . ($row + 1), $cellValue);
                    }
                    if (!empty($titleWidth)) {
                        $worksheet->getColumnDimension(WorkSheetHelper::cellName($_cols))->setWidth($titleWidth); // 列宽度
                    } else {
                        $worksheet->getColumnDimension(WorkSheetHelper::cellName($_cols))->setAutoSize(true); // 自动计算宽度
                    }
                    if ($num < count($val)) {
                        $col++;
                        $num++;
                    }
                    $_cols++;
                }
                $worksheet->mergeCells($_merge . $row . ':' . WorkSheetHelper::cellName($col) . $row);
                if (!isset($titleConfig['title_show']) || $titleConfig['title_show'] !== false) {
                    // 检查主标题是否必填（根据第一个子字段判断）
                    $firstSubField = reset($val);
                    $cellValue = self::formatHeaderTitle($key, $firstSubField, $requiredFields);
                    $worksheet->setCellValue($_merge . $row, $cellValue);
                }
            } else {
                if ($titleRow != 1) {
                    $worksheet->mergeCells($rowName . $row . ':' . $rowName . ($row + $titleRow - 1));
                }
                if (!isset($titleConfig['title_show']) || $titleConfig['title_show'] !== false) {
                    // 检查字段是否必填并格式化标题
                    $cellValue = self::formatHeaderTitle($key, $val, $requiredFields);
                    $worksheet->setCellValue($rowName . $row, $cellValue);
                }
                if (!empty($titleWidth)) {
                    $worksheet->getColumnDimension($rowName)->setWidth($titleWidth); // 列宽度
                } else {
                    $worksheet->getColumnDimension($rowName)->setAutoSize(true); // 自动计算宽度
                }
            }
            $col++;
            $_merge = WorkSheetHelper::cellName($col);
        }
        $row += $titleRow; // 当前行数
        
        return ['col' => $col, 'row' => $row];
    }

    /**
     * 格式化表头标题，如果字段必填则添加红色星号
     * @param string $title 表头标题
     * @param string|array $fieldName 字段名
     * @param array $requiredFields 必填字段列表
     * @return string|RichText 如果必填返回RichText对象，否则返回字符串
     */
    private static function formatHeaderTitle(string $title, $fieldName, array $requiredFields) {
        // 处理字段名为数组的情况
        if (is_array($fieldName)) {
            $fieldName = reset($fieldName);
        }
        
        // 检查字段是否必填
        if (!empty($requiredFields) && in_array($fieldName, $requiredFields)) {
            // 创建富文本对象
            $richText = new RichText();
            // 添加红色星号
            $redAsterisk = $richText->createTextRun('*');
            $redAsterisk->getFont()->setColor(new Color(Color::COLOR_RED));
            // 添加标题文本
            $richText->createText($title);
            return $richText;
        }
        
        // 非必填字段直接返回标题
        return $title;
    }
}


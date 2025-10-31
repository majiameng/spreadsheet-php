<?php
/**
 * @name: CellValueHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: CellValueHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel\Handler;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use tinymeng\spreadsheet\Util\WorkSheetHelper;

class CellValueHandler
{
    /**
     * 设置单元格值
     * @param Worksheet $worksheet
     * @param array $val 数据行
     * @param array $fields 字段列表
     * @param int $row 当前行
     * @param int $titleRow 标题行数
     * @param int|null $height 行高
     * @param bool $autoDataType 是否自动数据类型
     * @param bool $format 是否格式化内容
     * @param string $formatDate 日期格式
     * @return int 返回更新后的行号
     */
    public static function setCellValue(
        Worksheet $worksheet,
        array $val,
        array $fields,
        int $row,
        int $titleRow,
        ?int $height = null,
        bool $autoDataType = false,
        bool $format = true,
        string $formatDate = 'Y-m-d H:i:s'
    ): int {
        // 设置单元格行高
        if (!empty($height)) {
            $worksheet->getRowDimension($row)->setRowHeight($height);
        }
        
        $_lie = 0;
        foreach ($fields as $v) {
            $rowName = WorkSheetHelper::cellName($_lie);

            // 处理嵌套字段（如 'user.name'）
            if (strpos($v, '.') !== false) {
                $v = explode('.', $v);
                $content = $val;
                for ($i = 0; $i < count($v); $i++) {
                    $content = $content[$v[$i]] ?? '';
                }
            } elseif ($v == '_id') {
                $content = $row - $titleRow; // 自增序号列
            } else {
                $content = ($val[$v] ?? '');
            }

            // 处理图片类型
            if (is_array($content) && isset($content['type']) && isset($content['content'])) {
                if ($content['type'] == 'image') {
                    self::setImage($worksheet, $content, $rowName, $row);
                }
            }
            // 处理公式类型
            elseif (is_array($content) && isset($content['formula'])) {
                $worksheet->setCellValueExplicit(
                    $rowName . $row,
                    $content['formula'],
                    DataType::TYPE_FORMULA
                );
            }
            // 处理普通值
            else {
                $content = WorkSheetHelper::formatValue($content, $format, $formatDate); // 格式化数据
                if (is_numeric($content)) {
                    if ($autoDataType && strlen($content) < 11) {
                        $worksheet->setCellValueExplicit($rowName . $row, $content, DataType::TYPE_NUMERIC);
                    } else {
                        $worksheet->setCellValueExplicit($rowName . $row, $content, DataType::TYPE_STRING2);
                    }
                } else {
                    $worksheet->setCellValueExplicit($rowName . $row, $content, DataType::TYPE_STRING2);
                }
            }
            $_lie++;
        }
        $row++;
        
        return $row;
    }

    /**
     * 设置图片到单元格
     * @param Worksheet $worksheet
     * @param array $imageConfig 图片配置 ['type'=>'image', 'content'=>'路径', 'height'=>100, 'width'=>100, 'offsetX'=>0, 'offsetY'=>0]
     * @param string $rowName 列字母
     * @param int $row 行号
     */
    private static function setImage(
        Worksheet $worksheet,
        array $imageConfig,
        string $rowName,
        int $row
    ) {
        $path = WorkSheetHelper::verifyFile($imageConfig['content']);
        $drawing = new Drawing();
        $drawing->setPath($path);
        
        if (!empty($imageConfig['height'])) {
            $drawing->setHeight($imageConfig['height']);
        }
        if (!empty($imageConfig['width'])) {
            $drawing->setWidth($imageConfig['width']); // 只设置高，宽会自适应，如果设置宽后，高则失效
        }
        if (!empty($imageConfig['offsetX'])) {
            $drawing->setOffsetX($imageConfig['offsetX']); // 设置X方向偏移量
        }
        if (!empty($imageConfig['offsetY'])) {
            $drawing->setOffsetY($imageConfig['offsetY']); // 设置Y方向偏移量
        }

        $drawing->setCoordinates($rowName . $row);
        $drawing->setWorksheet($worksheet);
    }
}


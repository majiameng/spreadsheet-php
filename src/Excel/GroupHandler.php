<?php
/**
 * @name: GroupHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: GroupHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use tinymeng\tools\exception\StatusCode;
use tinymeng\tools\exception\TinymengException;

class GroupHandler
{
    /**
     * 处理左侧分组数据
     * @param Worksheet $worksheet
     * @param array $data 分组后的数据
     * @param int $groupLeftCount 分组级别数
     * @param array $groupLeft 分组字段列表
     * @param array $fields 所有字段列表
     * @param array $mergeColumns 需要合并的列
     * @param int $row 当前行
     * @param callable $cellNameFunc 获取列字母的函数
     * @param callable $setCellValueFunc 设置单元格值的函数
     * @return int 返回更新后的行号
     */
    public static function processGroupLeft(
        Worksheet $worksheet,
        array $data,
        int $groupLeftCount,
        array $groupLeft,
        array $fields,
        array $mergeColumns,
        int $row,
        callable $cellNameFunc,
        callable $setCellValueFunc
    ): int {
        // 获取分组字段在field中的实际位置
        $group_field_positions = [];
        foreach ($groupLeft as $group_field) {
            $position = array_search($group_field, $fields);
            if ($position !== false) {
                $group_field_positions[] = $position;
            }
        }

        if (empty($group_field_positions)) {
            throw new TinymengException(StatusCode::COMMON_PARAM_INVALID, '分组字段未在标题中定义');
        }

        $group_start = $row;
        foreach ($data as $key => $val) {
            // 第一级分组的合并单元格
            $rowName = $cellNameFunc($group_field_positions[0]); // 使用第一个分组字段的实际位置
            $coordinate = $rowName . $row . ':' . $rowName . ($row + $val['count'] - 1);
            $worksheet->mergeCells($coordinate);
            $worksheet->setCellValue($rowName . $row, $key);

            // 合并mergeColumns指定的其它列
            if (!empty($mergeColumns)) {
                foreach ($mergeColumns as $field) {
                    // 跳过分组字段本身
                    if (in_array($field, $groupLeft)) continue;
                    $colIdx = array_search($field, $fields);
                    if ($colIdx !== false) {
                        $colLetter = $cellNameFunc($colIdx);
                        $worksheet->mergeCells($colLetter . $row . ':' . $colLetter . ($row + $val['count'] - 1));
                        // 取本组第一个数据的值
                        $worksheet->setCellValue($colLetter . $row, $val['data'][0][$field] ?? '');
                    }
                }
            }

            if ($groupLeftCount == 1) {
                foreach ($val['data'] as $dataRow) {
                    $setCellValueFunc($dataRow);
                }
            } else {
                $sub_group_start = $row;
                $rowName = $cellNameFunc($group_field_positions[1]); // 使用第二个分组字段的实际位置

                foreach ($val['data'] as $k => $v) {
                    $coordinate = $rowName . $sub_group_start . ':' . $rowName . ($sub_group_start + $v['count'] - 1);
                    $worksheet->mergeCells($coordinate);
                    $worksheet->setCellValue($rowName . $sub_group_start, $k);

                    foreach ($v['data'] as $data) {
                        $setCellValueFunc($data);
                    }

                    $sub_group_start = $sub_group_start + $v['count'];
                }
            }

            $row = $group_start + $val['count'];
            $group_start = $row;
        }
        
        return $row;
    }

    /**
     * 数据分组（一级分组）
     * @param array $data 原始数据
     * @param string $groupField 分组字段
     * @return array
     */
    public static function groupDataByOneField(array $data, string $groupField): array
    {
        $grouped = [];
        foreach ($data as $k => $v) {
            if (isset($v[$groupField])) {
                $grouped[$v[$groupField]][] = $v;
            }
        }
        foreach ($grouped as $k => $v) {
            $grouped[$k] = [
                'data' => $v,
                'count' => count($v)
            ];
        }
        return $grouped;
    }

    /**
     * 数据分组（二级分组）
     * @param array $data 原始数据
     * @param string $firstGroupField 第一级分组字段
     * @param string $secondGroupField 第二级分组字段
     * @return array
     */
    public static function groupDataByTwoFields(array $data, string $firstGroupField, string $secondGroupField): array
    {
        $grouped = [];
        foreach ($data as $v) {
            if (isset($v[$firstGroupField]) && isset($v[$secondGroupField])) {
                $grouped[$v[$firstGroupField]][$v[$secondGroupField]][] = $v;
            }
        }
        return self::arrayCount($grouped);
    }

    /**
     * 二位数组获取每一级别数量
     * @param array $data 二维数组原始数据
     * @return array
     */
    private static function arrayCount(array $data): array
    {
        foreach ($data as $key => $val) {
            $num = 0;
            foreach ($val as $k => $v) {
                $sub_num = count($v);
                $num = $num + $sub_num;
                $val[$k] = [
                    'count' => $sub_num,
                    'data' => $v
                ];
            }
            $data[$key] = [
                'count' => $num,
                'data' => $val
            ];
        }
        return $data;
    }
}


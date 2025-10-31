<?php
/**
 * @name: DataValidationHandler
 * @author: JiaMeng <666@majiameng.com>
 * @file: DataValidationHandler.php
 * @Date: 2025/01/XX
 */
namespace tinymeng\spreadsheet\Excel\Handler;

use PhpOffice\PhpSpreadsheet\Cell\DataValidation;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use tinymeng\spreadsheet\Util\WorkSheetHelper;

class DataValidationHandler
{
    /**
     * 应用数据验证到指定列
     * @param Worksheet $worksheet 工作表对象
     * @param string $fieldName 字段名
     * @param array $config 验证配置
     * @param int $fieldIndex 字段在field数组中的索引
     * @param int $startRow 起始行（数据开始行）
     * @param int $endRow 结束行（数据结束行，为0时表示应用到整列）
     * @param bool $isTemplate 是否为模板（无数据）
     */
    public static function applyValidation(
        Worksheet $worksheet,
        string $fieldName,
        array $config,
        int $fieldIndex,
        int $startRow,
        int $endRow = 0,
        bool $isTemplate = false
    ) {
        $colLetter = WorkSheetHelper::cellName($fieldIndex);

        // 创建数据验证对象
        $validation = new DataValidation();
        
        // 设置验证类型
        self::setValidationType($validation, $config);
        
        // 设置操作符和范围（对于数值、日期、时间类型）
        self::setValidationOperator($validation, $config);
        
        // 设置输入提示信息
        self::setPromptMessage($validation, $config);
        
        // 设置错误提示信息
        self::setErrorMessage($validation, $config);
        
        // 是否允许空白
        $validation->setAllowBlank(isset($config['allowBlank']) ? $config['allowBlank'] : false);

        // 确定结束行：优先使用配置中的行范围
        $finalEndRow = self::calculateEndRow($config, $startRow, $endRow, $isTemplate);

        // 应用验证到指定范围
        $cellRange = self::calculateCellRange($colLetter, $startRow, $finalEndRow, $worksheet);
        $worksheet->setDataValidation($cellRange, $validation);
    }

    /**
     * 设置验证类型
     * @param DataValidation $validation
     * @param array $config
     */
    private static function setValidationType(DataValidation $validation, array $config)
    {
        $type = $config['type'] ?? 'list';
        switch ($type) {
            case 'list':
                $validation->setType(DataValidation::TYPE_LIST);
                if (isset($config['options']) && is_array($config['options'])) {
                    // 选项列表，使用逗号分隔，需要转义包含逗号的选项
                    $options = array_map(function($option) {
                        // 如果选项包含逗号或引号，需要用引号包裹并转义内部引号
                        if (strpos($option, ',') !== false || strpos($option, '"') !== false) {
                            return '"' . str_replace('"', '""', $option) . '"';
                        }
                        return $option;
                    }, $config['options']);
                    $formula = '"' . implode(',', $options) . '"';
                    $validation->setFormula1($formula);
                } elseif (isset($config['formula'])) {
                    // 使用公式引用范围（如 "=$A$1:$A$10"）
                    $validation->setFormula1($config['formula']);
                }
                // 是否显示下拉箭头
                $validation->setShowDropDown(!isset($config['showDropDown']) || $config['showDropDown'] !== false);
                break;
            case 'whole':
                $validation->setType(DataValidation::TYPE_WHOLE);
                break;
            case 'decimal':
                $validation->setType(DataValidation::TYPE_DECIMAL);
                break;
            case 'date':
                $validation->setType(DataValidation::TYPE_DATE);
                break;
            case 'time':
                $validation->setType(DataValidation::TYPE_TIME);
                break;
            case 'textLength':
                $validation->setType(DataValidation::TYPE_TEXTLENGTH);
                break;
            case 'custom':
                $validation->setType(DataValidation::TYPE_CUSTOM);
                if (isset($config['formula'])) {
                    $validation->setFormula1($config['formula']);
                }
                break;
            default:
                $validation->setType(DataValidation::TYPE_NONE);
        }
    }

    /**
     * 设置验证操作符和范围
     * @param DataValidation $validation
     * @param array $config
     */
    private static function setValidationOperator(DataValidation $validation, array $config)
    {
        $type = $config['type'] ?? 'list';
        
        if (!in_array($type, ['whole', 'decimal', 'date', 'time', 'textLength'])) {
            return;
        }

        $operator = $config['operator'] ?? 'between';
        switch ($operator) {
            case 'between':
                $validation->setOperator(DataValidation::OPERATOR_BETWEEN);
                if (isset($config['min'])) {
                    $validation->setFormula1($config['min']);
                }
                if (isset($config['max'])) {
                    $validation->setFormula2($config['max']);
                }
                break;
            case 'notBetween':
                $validation->setOperator(DataValidation::OPERATOR_NOTBETWEEN);
                if (isset($config['min'])) {
                    $validation->setFormula1($config['min']);
                }
                if (isset($config['max'])) {
                    $validation->setFormula2($config['max']);
                }
                break;
            case 'equal':
                $validation->setOperator(DataValidation::OPERATOR_EQUAL);
                if (isset($config['value'])) {
                    $validation->setFormula1($config['value']);
                }
                break;
            case 'notEqual':
                $validation->setOperator(DataValidation::OPERATOR_NOTEQUAL);
                if (isset($config['value'])) {
                    $validation->setFormula1($config['value']);
                }
                break;
            case 'greaterThan':
                $validation->setOperator(DataValidation::OPERATOR_GREATERTHAN);
                if (isset($config['value'])) {
                    $validation->setFormula1($config['value']);
                }
                break;
            case 'lessThan':
                $validation->setOperator(DataValidation::OPERATOR_LESSTHAN);
                if (isset($config['value'])) {
                    $validation->setFormula1($config['value']);
                }
                break;
            case 'greaterThanOrEqual':
                $validation->setOperator(DataValidation::OPERATOR_GREATERTHANOREQUAL);
                if (isset($config['value'])) {
                    $validation->setFormula1($config['value']);
                }
                break;
            case 'lessThanOrEqual':
                $validation->setOperator(DataValidation::OPERATOR_LESSTHANOREQUAL);
                if (isset($config['value'])) {
                    $validation->setFormula1($config['value']);
                }
                break;
        }
    }

    /**
     * 设置输入提示信息
     * @param DataValidation $validation
     * @param array $config
     */
    private static function setPromptMessage(DataValidation $validation, array $config)
    {
        if (isset($config['promptTitle']) || isset($config['promptMessage'])) {
            $validation->setPromptTitle($config['promptTitle'] ?? '');
            $validation->setPrompt($config['promptMessage'] ?? '');
            $validation->setShowInputMessage(isset($config['showInputMessage']) ? $config['showInputMessage'] : true);
        }
    }

    /**
     * 设置错误提示信息
     * @param DataValidation $validation
     * @param array $config
     */
    private static function setErrorMessage(DataValidation $validation, array $config)
    {
        if (isset($config['errorTitle']) || isset($config['errorMessage'])) {
            $validation->setErrorTitle($config['errorTitle'] ?? '输入错误');
            $validation->setError($config['errorMessage'] ?? '输入值无效');
            
            // 设置错误样式
            $errorStyle = $config['errorStyle'] ?? 'stop';
            switch ($errorStyle) {
                case 'stop':
                    $validation->setErrorStyle(DataValidation::STYLE_STOP);
                    break;
                case 'warning':
                    $validation->setErrorStyle(DataValidation::STYLE_WARNING);
                    break;
                case 'information':
                    $validation->setErrorStyle(DataValidation::STYLE_INFORMATION);
                    break;
                default:
                    $validation->setErrorStyle(DataValidation::STYLE_STOP);
            }
            
            $validation->setShowErrorMessage(isset($config['showErrorMessage']) ? $config['showErrorMessage'] : true);
        }
    }

    /**
     * 计算结束行
     * @param array $config
     * @param int $startRow
     * @param int $endRow
     * @param bool $isTemplate
     * @return int
     */
    private static function calculateEndRow(array $config, int $startRow, int $endRow, bool $isTemplate): int
    {
        $finalEndRow = $endRow;
        if (isset($config['data_end_row']) && $config['data_end_row'] > 0) {
            // 配置中直接指定了结束行
            $finalEndRow = $config['data_end_row'];
        } elseif (isset($config['data_row_count']) && $config['data_row_count'] > 0) {
            // 配置中指定了行数（从起始行开始计算）
            $finalEndRow = $startRow + $config['data_row_count'];
        } elseif ($endRow == 0 && $isTemplate) {
            // 模板导出时，如果没有配置行范围，默认应用到后续100行
            $finalEndRow = $startRow + 100;
        }
        return $finalEndRow;
    }

    /**
     * 计算单元格范围
     * @param string $colLetter
     * @param int $startRow
     * @param int $finalEndRow
     * @param Worksheet $worksheet
     * @return string
     */
    private static function calculateCellRange(string $colLetter, int $startRow, int $finalEndRow, Worksheet $worksheet): string
    {
        if ($finalEndRow > 0 && $finalEndRow >= $startRow) {
            // 应用到指定行范围
            return $colLetter . $startRow . ':' . $colLetter . $finalEndRow;
        } else {
            // 应用到整列（从数据开始行到工作表最后一行）
            return $colLetter . $startRow . ':' . $colLetter . $worksheet->getHighestRow();
        }
    }
}


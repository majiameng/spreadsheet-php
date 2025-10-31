<?php
/**
 * @name: WorkSheetHelper
 * @author: JiaMeng <666@majiameng.com>
 * @file: WorkSheetHelper.php
 * @Date: 2025/01/XX
 * @description: Excel工作表辅助工具类
 */
namespace tinymeng\spreadsheet\Util;

class WorkSheetHelper
{
    /**
     * 数字转英文列
     * @param int $columnIndex 列索引（从0开始）
     * @return string 列字母（如 A, B, C, AA, AB...）
     */
    public static function cellName(int $columnIndex): string
    {
        $columnIndex = (int)$columnIndex + 1;
        static $indexCache = [];

        if (!isset($indexCache[$columnIndex])) {
            $indexValue = $columnIndex;
            $base26 = null;
            do {
                $characterValue = ($indexValue % 26) ?: 26;
                $indexValue = ($indexValue - $characterValue) / 26;
                $base26 = chr($characterValue + 64) . ($base26 ?: '');
            } while ($indexValue > 0);
            $indexCache[$columnIndex] = $base26;
        }

        return $indexCache[$columnIndex];
    }

    /**
     * 格式化值
     * @param mixed $value 原始值
     * @param bool $format 是否格式化
     * @param string $formatDate 日期格式
     * @return mixed 格式化后的值
     */
    public static function formatValue($value, bool $format = true, string $formatDate = 'Y-m-d H:i:s')
    {
        if ($format === false) {
            return $value;
        }
        if (is_numeric($value) && strlen($value) === 10) {
            $value = date($formatDate, $value); // 时间戳转时间格式
        }
        return $value;
    }

    /**
     * 验证并处理文件路径
     * @param string $path 文件路径
     * @param bool $verifyFile 是否验证文件
     * @param \ZipArchive|null $zip ZIP归档对象
     * @return string 验证后的文件路径
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public static function verifyFile(string $path, bool $verifyFile = true, $zip = null): string
    {
        if ($verifyFile && preg_match('~^data:image/[a-z]+;base64,~', $path) !== 1) {
            // Check if a URL has been passed. https://stackoverflow.com/a/2058596/1252979
            if (filter_var($path, FILTER_VALIDATE_URL)) {
                $imageContents = file_get_contents($path);
                $filePath = tempnam(sys_get_temp_dir(), 'Drawing');
                if ($filePath) {
                    file_put_contents($filePath, $imageContents);
                    if (file_exists($filePath)) {
                        return $filePath;
                    }
                }
            } elseif (file_exists($path)) {
                return $path;
            } elseif ($zip instanceof \ZipArchive) {
                $zipPath = explode('#', $path)[1];
                if ($zip->locateName($zipPath) !== false) {
                    return $path;
                }
            } else {
                throw new \PhpOffice\PhpSpreadsheet\Exception("File $path not found!");
            }
        }
        return $path;
    }
}


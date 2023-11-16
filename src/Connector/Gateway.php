<?php
namespace tinymeng\spreadsheet\Connector;

use PhpOffice\PhpSpreadsheet\Exception as PhpSpreadsheetException;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Gateway
 */
abstract class Gateway implements GatewayInterface
{

    /**
     * @var Spreadsheet
     */
    public $spreadSheet;
    /**
     * @var Worksheet
     */
    public $workSheet;

    /**
     * 是否格式化内容
     * @var string
     */
    public $format = true;
    /**
     * 是否格式化内容
     * @var string
     */
    public $format_date = 'Y-m-d H:i:s';

    /**
     * 数字转英文列
     * @param $columnIndex
     * @return string
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/4/24 17:35
     */
    protected function cellName($columnIndex){
        $columnIndex =(int)$columnIndex+1;
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
     * 格式化value
     * @param string
     * @return mixed
     */
    protected function formatValue($v){
        if($this->format === false) return $v;
        if(is_numeric($v) && strlen($v)===10){
            $v = date($this->format_date,$v);//时间戳转时间格式
        }elseif (is_numeric($v) && strlen($v)>=19){
            $v = ' '.$v;//长数字在excel中会变科学计数法
        }
        return $v;
    }

    /**
     * 根据最后一列获取所有列数组
     * @param $lastCell
     * @return array
     */
    protected function getCellName($lastCell){
        $cellName = array();
        for($i='A'; $i!=$lastCell; $i++) {
            $cellName[] = $i;
        }
        $cellName[] = $i++;
        return $cellName;
    }

    /**
     * @param $url
     * @param $path
     * @return array|string|string[]
     */
    protected function verifyFile($path, $verifyFile = true, $zip = null){
        if ($verifyFile && preg_match('~^data:image/[a-z]+;base64,~', $path) !== 1) {
            // Check if a URL has been passed. https://stackoverflow.com/a/2058596/1252979
            if (filter_var($path, FILTER_VALIDATE_URL)) {
                $this->path = $path;
                // Implicit that it is a URL, rather store info than running check above on value in other places.
                $this->isUrl = true;
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
            } elseif ($zip instanceof ZipArchive) {
                $zipPath = explode('#', $path)[1];
                if ($zip->locateName($zipPath) !== false) {
                    return $path;
                }
            } else {
                throw new PhpSpreadsheetException("File $path not found!");
            }
        }
        return $path;
    }
}

<?php
namespace tinymeng\spreadsheet\Connector;

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
     * numToEn
     * @param $columnIndex
     * @return string
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/4/24 17:35
     */
    public function cellName($columnIndex){
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
    public function formatValue($v){
        if($this->format === false) return $v;
        if(is_numeric($v) && strlen($v)===10){
            $v = date($this->format_date,$v);//时间戳转时间格式
        }elseif (is_numeric($v) && strlen($v)>=19){
            $v = ' '.$v;//长数字在excel中会变科学计数法
        }
        return $v;
    }
}

<?php
/**
 * @name: SpreadSheet
 * @author: JiaMeng <666@majiameng.com>
 * @file: Export.php
 * @Date: 2024/03/04 10:15
 */
namespace tinymeng\spreadsheet\Excel;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use tinymeng\tools\FileTool;

trait SpreadSheet{

    /**
     * sheet名称
     * @var
     */
    private $sheetName;
    /**
     * 定义所有字段
     * @var array
     */
    private $field = [];

    /**
     * 文件信息
     * @var array
     */
    private $fileTitle=[];

    /**
     * 标题占用行数
     * @var int
     */
    private $title_row = 1;

    /**
     * 左侧分组字段
     * @var array
     */
    private $group_left = [];


    /**
     * 表头所在行
     * @var int
     */
    public $titleFieldsRow = 1;

    /**
     * 获取表格列数
     * @var
     */
    public $columnCount;

    /**
     * 获取表格行数
     * @var
     */
    public $rowCount;
    /**
     * title
     * @var
     */
    public $title;
    /**
     * title字段
     * @var
     */
    public $title_fields;
    /**
     * @var string[]
     */
    private $cellName = [];

    /**
     * 文件中图片读取
     * 图片存储的相对路径
     * @var string
     */
    public $relative_path = '/images';

    /**
     * 文件中图片读取
     * 图片存储的绝对路径
     * @var string
     */
    public $image_path = '/images';

    public function setTitle($title){
        $this->title = $title;
        $this->getTitleFields();
        return $this;
    }

    /**
     * @param $value
     * @return $this
     */
    public function setRelativePath($value){
        $this->relative_path = $value;
        return $this;
    }

    /**
     * @param $value
     * @return $this
     */
    public function setImagePath($value){
        $this->image_path = $value;
        return $this;
    }

    /**
     * getExcelData
     * @return array
     * @throws Exception
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:30
     */
    public function getExcelData(){
        /* 循环读取每个单元格的数据 */
        $result = [];
        $dataRow = $this->titleFieldsRow+1;

        //行数循环
        for ($row = $dataRow; $row <= $this->rowCount; $row++){
            $rowFlog = false;//行是否有内容（过滤空行）
            //列数循环 , 列数是以A列开始
            $data = [];
            foreach ($this->cellName as $column){
                $cell = $this->workSheet->getCell($column.$row);
                $value = trim($cell->getFormattedValue());
                if(isset($this->title_fields[$column])){
                    $data[$this->title_fields[$column]] = $value;
                    if(!empty($value)) $rowFlog = true;//有内容
                }
            }
            if($rowFlog) $result[] = $data;
        }

        /*
         * 读取表格图片数据
         * (如果为空右击图片转为浮动图片)
         */
        $imageFilename_prefix = time().rand(100,999).$this->sheet;
        foreach ($this->workSheet->getDrawingCollection() as $drawing) {
            /**@var $drawing Drawing* */
            list($column, $row) = Coordinate::coordinateFromString($drawing->getCoordinates());
            $imageFilename = "/{$imageFilename_prefix}-" . $drawing->getCoordinates();
            $image_suffix = $this->saveImage($drawing, $imageFilename);
            $image_name = ltrim($this->relative_path, '/') . "{$imageFilename}.{$image_suffix}";
            if(isset($this->title_fields[$column])) {
                $result[$row-($this->titleFieldsRow+1)][$this->title_fields[$column]] = $image_name;
            }
        }
        return $result;
    }

    /**
     * getTitle
     * @return mixed
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:30
     */
    public function getTitle(){
        return $this->title;
    }

    /**
     * getTitleFields
     * @return SpreadSheet
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:30
     */
    public function getTitleFields(){
        $title = $this->getTitle();

        $row = $this->titleFieldsRow;
        $titleDataArr = [];

        foreach ($this->cellName as $column){
            $value = trim($this->workSheet->getCell($column.$row)->getValue());
            // 规范化表头：移除前导星号（半角/全角）与多余空白
            $norm = $this->normalizeHeaderName($value);
            if(!empty($norm)){
                $titleDataArr[$norm] = $column;
            }
        }
        $title_fields = [];
        foreach ($title as $key=>$value) {
            $normKey = $this->normalizeHeaderName($key);
            if(isset($titleDataArr[$normKey])){
                $title_fields[$titleDataArr[$normKey]] = $value;
            }
        }
        $this->title_fields = $title_fields;
        return $this;
    }

    /**
     * 规范化表头显示名称：
     * - 去掉前导的 * 或 ＊，以及其后的空格
     * - 去掉首尾空白
     * @param $name
     * @return mixed|string
     */
    private function normalizeHeaderName($name){
        $name = is_string($name) ? trim($name) : $name;
        if(!is_string($name)) return $name;
        // 移除一个或多个半角/全角星号以及紧随的空格
        $name = preg_replace('/^[\x{002A}\x{FF0A}]+\s*/u', '', $name);
        // 再次trim以防有残余空白
        return trim($name);
    }

    /**
     * 保存图片到文件相对路径
     * @param Drawing $drawing
     * @param $imageFilename
     * @return string
     * @throws Exception
     */
    protected function saveImage(Drawing $drawing, $imageFilename)
    {
        FileTool::mkdir($this->image_path);

        // 获取文件的真实MIME类型
        $fInfo = new \finfo(FILEINFO_MIME_TYPE);
        $mimeType = $fInfo->file($drawing->getPath());

        // 根据MIME类型确定真实的图片格式
        switch ($mimeType) {
            case 'image/jpg':
            case 'image/jpeg':
                $realExtension = 'jpg';
                $imageFilename .= '.'.$realExtension;
                $source = imagecreatefromjpeg($drawing->getPath());
                imagejpeg($source, $this->image_path . $imageFilename);
                break;
            case 'image/gif':
                $realExtension = 'gif';
                $imageFilename .= '.'.$realExtension;
                $source = imagecreatefromgif($drawing->getPath());
                imagegif($source, $this->image_path . $imageFilename);
                break;
            case 'image/png':
                $realExtension = 'png';
                $imageFilename .= '.'.$realExtension;
                $source = imagecreatefrompng($drawing->getPath());
                // 保持透明度设置
                imagealphablending($source, false);
                imagesavealpha($source, true);
                imagepng($source, $this->image_path . $imageFilename);
                break;
            default:
                throw new Exception('image format error!');
        }

        return $realExtension;
    }
}

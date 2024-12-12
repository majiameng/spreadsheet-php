<?php
/**
 * @name: TWorklSheet
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
     * @param $this->title_fields
     * @return array
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
        $image_filename_prefix = time().rand(100,999).$this->sheet;
        foreach ($this->workSheet->getDrawingCollection() as $drawing) {
            /**@var $drawing Drawing* */
            list($column, $row) = Coordinate::coordinateFromString($drawing->getCoordinates());
            $image_filename = "/{$image_filename_prefix}-" . $drawing->getCoordinates();
            $image_suffix = $this->saveImage($drawing, $image_filename);
            $image_name = ltrim($this->relative_path, '/') . "{$image_filename}.{$image_suffix}";
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
     * @return array
     * @author: Tinymeng <666@majiameng.com>
     * @time: 2022/2/22 11:30
     */
    public function getTitleFields(){
        $title = $this->getTitle();

        $row = $this->titleFieldsRow;
        $titleDataArr = [];

        foreach ($this->cellName as $column){
            $value = trim($this->workSheet->getCell($column.$row)->getValue());
            if(!empty($value)){
                $titleDataArr[$value] = $column;
            }
        }
        $title_fields = [];
        foreach ($title as $key=>$value) {
            if(isset($titleDataArr[$key])){
                $title_fields[$titleDataArr[$key]] = $value;
            }
        }
        $this->title_fields = $title_fields;
        return $this;
    }

    /**
     * 保存图片到文件相对路径
     * @param Drawing $drawing
     * @param $image_filename
     * @return string
     * @throws Exception
     */
    protected function saveImage(Drawing $drawing, $image_filename)
    {
        FileTool::mkdir($this->image_path);
        $image_filename .= '.' . $drawing->getExtension();
        switch ($drawing->getExtension()) {
            case 'jpg':
            case 'jpeg':
                $source = imagecreatefromjpeg($drawing->getPath());
                imagejpeg($source, $this->image_path . $image_filename);
                break;
            case 'gif':
                $source = imagecreatefromgif($drawing->getPath());
                imagegif($source, $this->image_path . $image_filename);
                break;
            case 'png':
                $source = imagecreatefrompng($drawing->getPath());
                imagepng($source, $this->image_path . $image_filename);
                break;
            default:
                throw new Exception('image format error!');
        }

        return $drawing->getExtension();
    }
}

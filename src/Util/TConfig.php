<?php
/**
 * @name: TConfig
 * @author: JiaMeng <666@majiameng.com>
 * @file: Export.php
 * @Date: 2024/03/04 10:15
 */
namespace tinymeng\spreadsheet\Util;

trait TConfig{

    /**
     * 文件创建者
     * @var string 
     */
    private $creator = 'tinymeng';

    /**
     * 文件名称
     * @var
     */
    private $fileName;

    /**
     * 文件存储位置
     * @var string|bool
     */
    private $pathName = false;

    /**
     * 是否居中
     * @var string
     */
    private $horizontalCenter = true;


    /**
     * 定义表头行高
     * @var int 常用：22
     */
    private $titleHeight = null;
    
    /**
     * 定义表头列宽(未设置则自动计算宽度)
     * @var int 常用：20
     */
    private $titleWidth = null;
    
    /**
     * 定义数据行高
     * @var int 常用：22
     */
    private $height = null;
    

    /**
     * 自动筛选(是否开启)
     * @var bool
     */
    private $autoFilter = false;

    /**
     * 自动适应文本类型
     * @var bool
     */
    private $autoDataType = true;

    /**
     * 冻结窗格（要冻结的首行首列"B2"，false不开启）
     * @var string|bool
     */
    private $freezePane = false;

    /**
     * 当前选中sheet
     * @var string|bool
     */
    private $sheet = 0;

    /**
     * 标题占用行数
     * @var int
     */
    private $title_row = 1;
    /**
     * 报表名称(主标题)
     * @var
     */
    private $mainTitle = '';

    /**
     * 是否需要报表名称(主标题)
     * @var bool
     */
    private $mainTitleLine = false;

    /**
     * @return bool|int|string
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * @param bool|int|string $sheet
     */
    public function setSheet($sheet)
    {
        $this->sheet = $sheet;
        return $this;
    }


    public function getTitleRow(): int
    {
        return $this->title_row;
    }

    public function setTitleRow(int $title_row)
    {
        $this->title_row = $title_row;
        return $this;
    }

    public function getCreator(): string
    {
        return $this->creator;
    }

    public function setCreator(string $creator)
    {
        $this->creator = $creator;
        return $this;
    }

    /**
     * @return mixed
     */
    public function getFileName()
    {
        return $this->fileName;
    }

    /**
     * @param mixed $fileName
     */
    public function setFileName($fileName)
    {
        $this->fileName = $fileName;
        return $this;
    }

    /**
     * @return bool|string
     */
    public function getHorizontalCenter()
    {
        return $this->horizontalCenter;
    }

    /**
     * @param bool|string $horizontalCenter
     */
    public function setHorizontalCenter($horizontalCenter)
    {
        $this->horizontalCenter = $horizontalCenter;
        return $this;
    }

    public function getTitleHeight(): int
    {
        return $this->titleHeight;
    }

    public function setTitleHeight(int $titleHeight)
    {
        $this->titleHeight = $titleHeight;
        return $this;
    }

    public function getTitleWidth(): int
    {
        return $this->titleWidth;
    }

    public function setTitleWidth(int $titleWidth)
    {
        $this->titleWidth = $titleWidth;
        return $this;
    }

    public function getHeight(): int
    {
        return $this->height;
    }

    public function setHeight(int $height)
    {
        $this->height = $height;
        return $this;
    }

    public function isAutoFilter(): bool
    {
        return $this->autoFilter;
    }

    public function setAutoFilter(bool $autoFilter)
    {
        $this->autoFilter = $autoFilter;
        return $this;
    }

    public function isAutoDataType(): bool
    {
        return $this->autoDataType;
    }

    public function setAutoDataType(bool $autoDataType)
    {
        $this->autoDataType = $autoDataType;
        return $this;
    }

    /**
     * @return bool|string
     */
    public function getFreezePane()
    {
        return $this->freezePane;
    }

    /**
     * @param bool|string $freezePane
     */
    public function setFreezePane($freezePane)
    {
        $this->freezePane = $freezePane;
        return $this;
    }

    /**
     * @param bool|string $pathName
     */
    public function setPathName($pathName)
    {
        $this->pathName = $pathName;
        return $this;
    }

    /**
     * @param $pathName
     * @return bool|mixed|string
     */
    private function getPathName($pathName)
    {
        if(!empty($pathName)){
            return $pathName;
        }
        if($this->pathName){
            return $this->pathName;
        }
        $pathName = dirname( dirname(dirname(SPREADSHEET_ROOT_PATH))).DIRECTORY_SEPARATOR."public".DIRECTORY_SEPARATOR."export".DIRECTORY_SEPARATOR.date('Ymd').DIRECTORY_SEPARATOR;
        return $pathName;
    }

    public function isMainTitleLine(): bool
    {
        return $this->mainTitleLine;
    }

    public function setMainTitleLine(bool $mainTitleLine): void
    {
        $this->mainTitleLine = $mainTitleLine;
    }

    public function getMainTitle(): string
    {
        return $this->mainTitle;
    }

    public function setMainTitle(string $mainTitle): void
    {
        $this->mainTitle = $mainTitle;
    }

}

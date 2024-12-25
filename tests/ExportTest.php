<?php

use PHPUnit\Framework\TestCase;
use tinymeng\spreadsheet\Gateways\Export;

class ExportTest extends TestCase
{
    protected $export;

    protected function setUp(): void
    {
        parent::setUp();
        $this->export = new Export([
            'creator' => 'Test Creator',
            'autoFilter' => true,
            'horizontalCenter' => true,
        ]);
    }

    public function testCreateWorkSheet()
    {
        $this->export->createWorkSheet('TestSheet');
        $this->assertEquals('TestSheet', $this->export->workSheet->getTitle());
    }


    public function testSaveFile()
    {
        $pathName = __DIR__ . '/tmp/';
        $fileName = $this->export->save('TestSheet', $pathName);
        $this->assertTrue(file_exists($fileName));
        unlink($fileName);
        rmdir($pathName);
    }
}

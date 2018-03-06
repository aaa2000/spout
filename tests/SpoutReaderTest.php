<?php

namespace Port\Spout\Tests;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Port\Exception\ReaderException;
use Port\Spout\SpoutReader;

class SpoutReaderTest extends \PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        if (!extension_loaded('zip')) {
            $this->markTestSkipped();
        }
    }

    public function testCountWithoutHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($fileReader);
        $this->assertEquals(3, $reader->count());
    }

    public function testCountWithHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($fileReader, 0);
        $this->assertEquals(3, $reader->count());
    }

    public function testGetColumnHeadersWithHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($fileReader, 0);
        $this->assertEquals(['id', 'number', 'description'], $reader->getColumnHeaders());
    }

    public function testGetColumnHeadersWithoutHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($fileReader);
        $this->assertNull($reader->getColumnHeaders());
    }

    public function testIterate()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($fileReader, 0);
        foreach ($reader as $row) {
            $this->assertInternalType('array', $row);
            $this->assertEquals(['id', 'number', 'description'], array_keys($row));
        }
    }

    public function testIterateWithoutHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($fileReader);

        $this->assertSame(
            [
                [50, 123, 'Description'],
                [6, 456, 'Another description'],
                [7, 7890, 'Some more info'],
            ],
            iterator_to_array($reader, false)
        );
    }

    public function testIterateWithHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($fileReader, 0);

        $this->assertSame(
            [
                ['id' => 50, 'number' => 123, 'description' => 'Description'],
                ['id' => 6, 'number' => 456, 'description' => 'Another description'],
                ['id' => 7, 'number' => 7890, 'description' => 'Some more info'],
            ],
            iterator_to_array($reader, false)
        );
    }

    public function testMultiSheet()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_multi_sheet.xlsx');
        $sheet1reader = new SpoutReader($fileReader, null, 0);
        $this->assertEquals(3, $sheet1reader->count());

        $sheet2reader = new SpoutReader($fileReader, null, 1);
        $this->assertEquals(2, $sheet2reader->count());
    }

    public function testWithSheetIndexOutOfRangeShouldThrowException()
    {
        $this->setExpectedException(ReaderException::class);
        $headerRowNumber = 1;
        $activeSheetIndex = 100;
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_multi_sheet.xlsx');
        new SpoutReader($fileReader, $headerRowNumber, $activeSheetIndex);
    }

    public function testSeek()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($fileReader);
        $reader->seek(0);
        $this->assertEquals([50, 123, 'Description'], $reader->current());
        $reader->seek(1);
        $this->assertEquals([6, 456, 'Another description'], $reader->current());
    }

    public function testSeekWhenPositionIsNotValid()
    {
        $this->setExpectedException(\OutOfBoundsException::class);
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($fileReader);
        $reader->seek(1000);
    }

    public function testSeekWithHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($fileReader, 0);
        $reader->seek(0);
        $this->assertEquals(['id' => 50, 'number' => 123, 'description' => 'Description'], $reader->current());
        $reader->seek(1);
        $this->assertEquals(['id' => 6, 'number' => 456, 'description' => 'Another description'], $reader->current());
    }

    public function testSetHeaders()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($fileReader);
        $reader->setColumnHeaders(['product id', 'amount', 'info']);
        $reader->seek(0);
        $this->assertEquals(['product id', 'amount', 'info'], array_keys($reader->current()));
    }
}


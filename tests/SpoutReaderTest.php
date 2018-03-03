<?php

namespace Port\Spout\Tests;

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
        $file = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($file);
        $this->assertEquals(3, $reader->count());
    }

    public function testCountWithHeaders()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($file, 0);
        $this->assertEquals(3, $reader->count());
    }

    public function testGetColumnHeadersWithHeaders()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($file, 0);
        $this->assertEquals(['id', 'number', 'description'], $reader->getColumnHeaders());
    }

    public function testGetColumnHeadersWithoutHeaders()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($file);
        $this->assertNull($reader->getColumnHeaders());
    }

    public function testIterate()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($file, 0);
        foreach ($reader as $row) {
            $this->assertInternalType('array', $row);
            $this->assertEquals(['id', 'number', 'description'], array_keys($row));
        }
    }

    public function testIterateWithoutHeaders()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($file);

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
        $file = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($file, 0);

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
        $file = new \SplFileObject(__DIR__.'/fixtures/data_multi_sheet.xlsx');
        $sheet1reader = new SpoutReader($file, null, 0);
        $this->assertEquals(3, $sheet1reader->count());

        $sheet2reader = new SpoutReader($file, null, 1);
        $this->assertEquals(2, $sheet2reader->count());
    }

    public function testWithSheetIndexOutOfRangeShouldThrowException()
    {
        $this->setExpectedException(ReaderException::class);
        $headerRowNumber = 1;
        $activeSheetIndex = 100;
        $file = new \SplFileObject(__DIR__.'/fixtures/data_multi_sheet.xlsx');
        new SpoutReader($file, $headerRowNumber, $activeSheetIndex);
    }

    public function testSeek()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($file);
        $reader->seek(0);
        $this->assertEquals([50, 123, 'Description'], $reader->current());
        $reader->seek(1);
        $this->assertEquals([6, 456, 'Another description'], $reader->current());
    }

    public function testSeekWhenPositionIsNotValid()
    {
        $this->setExpectedException(\OutOfBoundsException::class);
        $file = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($file);
        $reader->seek(1000);
    }

    public function testSeekWithHeaders()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx');
        $reader = new SpoutReader($file, 0);
        $reader->seek(0);
        $this->assertEquals(['id' => 50, 'number' => 123, 'description' => 'Description'], $reader->current());
        $reader->seek(1);
        $this->assertEquals(['id' => 6, 'number' => 456, 'description' => 'Another description'], $reader->current());
    }

    public function testGetRow()
    {
        $file = new \SplFileObject(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        $reader = new SpoutReader($file);
        $reader->getRow(0);
        $this->assertEquals([50, 123, 'Description'], $reader->current());
        $reader->getRow(1);
        $this->assertEquals([6, 456, 'Another description'], $reader->current());
    }
}


<?php

namespace Port\Spout\Tests;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
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

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testCountWithoutHeaders($fileReader)
    {
        $reader = new SpoutReader($fileReader);
        $this->assertEquals(3, $reader->count());
    }

    public function dataNoColumnHeadersFileProvider()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.xlsx');
        yield ['xlsx' => $fileReader];

        $fileReader = ReaderFactory::create(Type::ODS);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.ods');
        yield ['ods' => $fileReader];

        $fileReader = ReaderFactory::create(Type::CSV);
        $fileReader->open(__DIR__.'/fixtures/data_no_column_headers.csv');
        yield ['csv' => $fileReader];
    }

    /**
     * @dataProvider dataColumnHeadersFileProvider
     */
    public function testCountWithHeaders(ReaderInterface $fileReader)
    {
        $reader = new SpoutReader($fileReader, 0);
        $this->assertEquals(3, $reader->count());
    }

    public function dataColumnHeadersFileProvider()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.xlsx');
        yield ['xlsx' => $fileReader];

        $fileReader = ReaderFactory::create(Type::ODS);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.ods');
        yield ['ods' => $fileReader];

        $fileReader = ReaderFactory::create(Type::CSV);
        $fileReader->open(__DIR__.'/fixtures/data_column_headers.csv');
        yield ['csv' => $fileReader];
    }

    /**
     * @dataProvider dataColumnHeadersFileProvider
     */
    public function testGetColumnHeadersWithHeaders($fileReader)
    {
        $reader = new SpoutReader($fileReader, 0);
        $this->assertEquals(['id', 'number', 'description'], $reader->getColumnHeaders());
    }

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testGetColumnHeadersWithoutHeaders($fileReader)
    {
        $reader = new SpoutReader($fileReader);
        $this->assertNull($reader->getColumnHeaders());
    }

    /**
     * @dataProvider dataColumnHeadersFileProvider
     */
    public function testIterate($fileReader)
    {
        $reader = new SpoutReader($fileReader, 0);
        foreach ($reader as $row) {
            $this->assertInternalType('array', $row);
            $this->assertEquals(['id', 'number', 'description'], array_keys($row));
        }
    }

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testIterateWithoutHeaders($fileReader)
    {
        $reader = new SpoutReader($fileReader);

        $this->assertEquals(
            [
                [50, 123, 'Description'],
                [6, 456, 'Another description'],
                [7, 7890, 'Some more info'],
            ],
            iterator_to_array($reader, false)
        );
    }

    /**
     * @dataProvider dataColumnHeadersFileProvider
     */
    public function testIterateWithHeaders($fileReader)
    {
        $reader = new SpoutReader($fileReader, 0);

        $this->assertEquals(
            [
                ['id' => 50, 'number' => 123, 'description' => 'Description'],
                ['id' => 6, 'number' => 456, 'description' => 'Another description'],
                ['id' => 7, 'number' => 7890, 'description' => 'Some more info'],
            ],
            iterator_to_array($reader, false)
        );
    }

    /**
     * @dataProvider dataMultiSheetFileProvider
     */
    public function testMultiSheet($fileReader)
    {
        $sheet1reader = new SpoutReader($fileReader, null, 0);
        $this->assertEquals(3, $sheet1reader->count());

        $sheet2reader = new SpoutReader($fileReader, null, 1);
        $this->assertEquals(2, $sheet2reader->count());
    }

    public function dataMultiSheetFileProvider()
    {
        $fileReader = ReaderFactory::create(Type::XLSX);
        $fileReader->open(__DIR__.'/fixtures/data_multi_sheet.xlsx');
        yield ['xlsx' => $fileReader];

        $fileReader = ReaderFactory::create(Type::ODS);
        $fileReader->open(__DIR__.'/fixtures/data_multi_sheet.ods');
        yield ['ods' => $fileReader];
    }


    /**
     * @dataProvider dataMultiSheetFileProvider
     */
    public function testWithSheetIndexOutOfRangeShouldThrowException($fileReader)
    {
        $this->setExpectedException(ReaderException::class);
        $headerRowNumber = 1;
        $activeSheetIndex = 100;
        new SpoutReader($fileReader, $headerRowNumber, $activeSheetIndex);
    }

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testSeek($fileReader)
    {
        $reader = new SpoutReader($fileReader);
        $reader->seek(0);
        $this->assertEquals([50, 123, 'Description'], $reader->current());
        $reader->seek(1);
        $this->assertEquals([6, 456, 'Another description'], $reader->current());
    }

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testSeekWithDescendingOrder($fileReader)
    {
        $reader = new SpoutReader($fileReader);
        $reader->seek(1);
        $this->assertEquals([6, 456, 'Another description'], $reader->current());
        $reader->seek(0);
        $this->assertEquals([50, 123, 'Description'], $reader->current());
    }

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testSeekWhenPositionIsNotValid($fileReader)
    {
        $this->setExpectedException(\OutOfBoundsException::class);
        $reader = new SpoutReader($fileReader);
        $reader->seek(1000);
    }

    /**
     * @dataProvider dataColumnHeadersFileProvider
     */
    public function testSeekWithHeaders($fileReader)
    {
        $reader = new SpoutReader($fileReader, 0);
        $reader->seek(0);
        $this->assertEquals(['id' => 50, 'number' => 123, 'description' => 'Description'], $reader->current());
        $reader->seek(1);
        $this->assertEquals(['id' => 6, 'number' => 456, 'description' => 'Another description'], $reader->current());
    }

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testSetHeaders($fileReader)
    {
        $reader = new SpoutReader($fileReader);
        $reader->setColumnHeaders(['product id', 'amount', 'info']);
        $reader->seek(0);
        $this->assertEquals(['product id', 'amount', 'info'], array_keys($reader->current()));
    }

    /**
     * @dataProvider dataNoColumnHeadersFileProvider
     */
    public function testReadMultipleTimesShouldRewindReader($fileReader)
    {
        $reader = new SpoutReader($fileReader);

        $i = 0;
        foreach ($reader as $row) {
            $i++;
        }
        foreach ($reader as $row) {
            $i++;
        }

        $this->assertGreaterThan(0, $i);
    }
}


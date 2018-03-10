<?php

namespace Port\Spout\Tests;

use Port\Exception\UnexpectedValueException;
use Port\Spout\SpoutReaderFactory;
use Port\Spout\SpoutReader;

class SpoutReaderFactoryTest extends \PHPUnit_Framework_TestCase
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
    public function testGetReaderWithNoHeaderLine($filePath)
    {
        $factory = new SpoutReaderFactory();
        $reader = $factory->getReader(new \SplFileObject($filePath));
        $this->assertInstanceOf(SpoutReader::class, $reader);
        $this->assertCount(3, $reader);
    }

    public function dataNoColumnHeadersFileProvider()
    {
        yield ['xlsx' => __DIR__.'/fixtures/data_no_column_headers.xlsx'];
        yield ['ods' => __DIR__.'/fixtures/data_no_column_headers.ods'];
        yield ['csv' => __DIR__.'/fixtures/data_no_column_headers.csv'];
    }

    /**
     * @dataProvider dataColumnHeadersFileProvider
     */
    public function testGetReaderWithHeaderLine($filePath)
    {
        $factory = new SpoutReaderFactory(0);
        $reader = $factory->getReader(new \SplFileObject($filePath));
        $this->assertCount(3, $reader);
    }

    public function dataColumnHeadersFileProvider()
    {
        yield ['xlsx' => __DIR__.'/fixtures/data_column_headers.xlsx'];
        yield ['ods' => __DIR__.'/fixtures/data_column_headers.ods'];
        yield ['csv' => __DIR__.'/fixtures/data_column_headers.csv'];
    }

    public function testGetReaderWithUnknownFormatShouldThrowAnException()
    {
        $this->setExpectedException(UnexpectedValueException::class);
        $factory = new SpoutReaderFactory(null, null, true);
        $unknownExtensionFile = tempnam(sys_get_temp_dir(), null);
        $factory->getReader(new \SplFileObject($unknownExtensionFile));
    }
}

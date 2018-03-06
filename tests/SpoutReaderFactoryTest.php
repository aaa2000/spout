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

    public function testGetReaderWithExcelFormat()
    {
        $factory = new SpoutReaderFactory();
        $reader = $factory->getReader(new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx'));
        $this->assertInstanceOf(SpoutReader::class, $reader);
        $this->assertCount(4, $reader);
    }

    public function testGetReaderWithOdsFormat()
    {
        $this->markTestSkipped('The ods iterator can\'t be rewind more than once');
        $factory = new SpoutReaderFactory();
        $reader = $factory->getReader(new \SplFileObject(__DIR__.'/fixtures/data_column_headers.ods'));
        $this->assertInstanceOf(SpoutReader::class, $reader);
        $this->assertCount(4, $reader);
    }

    public function testGetReaderWithCsvFormat()
    {
        $factory = new SpoutReaderFactory();
        $reader = $factory->getReader(new \SplFileObject(__DIR__.'/fixtures/data_column_headers.csv'));
        $this->assertInstanceOf(SpoutReader::class, $reader);
        $this->assertCount(4, $reader);
    }

    public function testGetReaderWithHeaderLine()
    {
        $factory = new SpoutReaderFactory(0);
        $reader = $factory->getReader(new \SplFileObject(__DIR__ . '/fixtures/data_column_headers.xlsx'));
        $this->assertCount(3, $reader);
    }

    public function testGetReaderWithUnknownFormatShouldThrowAnException()
    {
        $this->setExpectedException(UnexpectedValueException::class);
        $factory = new SpoutReaderFactory(null, null, true);
        $unknownExtensionFile = tempnam(sys_get_temp_dir(), null);
        $factory->getReader(new \SplFileObject($unknownExtensionFile));
    }
}

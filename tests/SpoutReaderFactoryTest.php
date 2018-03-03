<?php

namespace Port\Spout\Tests;

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

    public function testGetReader()
    {
        $factory = new SpoutReaderFactory();
        $reader = $factory->getReader(new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx'));
        $this->assertInstanceOf(SpoutReader::class, $reader);
        $this->assertCount(4, $reader);

        $factory = new SpoutReaderFactory(0);
        $reader = $factory->getReader(new \SplFileObject(__DIR__.'/fixtures/data_column_headers.xlsx'));
        $this->assertCount(3, $reader);
    }
}

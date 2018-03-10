<?php

namespace Port\Spout\Tests;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\ReaderInterface;
use Box\Spout\Reader\SheetInterface;
use Box\Spout\Writer\WriterFactory;
use Port\Spout\SpoutReader;
use Port\Spout\SpoutWriter;

class SpoutWriterTest extends \PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        if (!extension_loaded('zip')) {
            $this->markTestSkipped();
        }
    }

    public function writerTypeProvider()
    {
        yield [Type::XLSX];
        yield [Type::ODS];
        yield [Type::CSV];
    }

    public function sheetWriterTypeProvider()
    {
        yield [Type::XLSX];
        yield [Type::ODS];
    }

    /**
     * @dataProvider sheetWriterTypeProvider
     */
    public function testWriteItemAppendWithSheetTitle($type)
    {
        $file = tempnam(sys_get_temp_dir(), null);
        $writerFile = WriterFactory::create($type);
        $writerFile->openToFile($file);
        $writer = new SpoutWriter($writerFile, 'Sheet 1');

        $writer->prepare();
        $writer->writeItem(['first', 'last']);

        $writer->writeItem([
            'first' => 'James',
            'last'  => 'Bond'
        ]);

        $writer->writeItem([
            'first' => '',
            'last'  => 'Dr. No'
        ]);

        $writer->finish();

        $reader = ReaderFactory::create($type);
        $reader->open($file);
        $spoutReader = new SpoutReader($reader);

        $this->assertCount(3, $spoutReader);
        $this->assertInstanceOf(SheetInterface::class, $this->getSheetByTitle($reader, 'Sheet 1'));
    }

    /**
     * @dataProvider writerTypeProvider
     */
    public function testWriteItemWithoutSheetTitle($type)
    {
        $file = tempnam(sys_get_temp_dir(), null);
        $writerFile = WriterFactory::create($type);
        $writerFile->openToFile($file);
        $writer = new SpoutWriter($writerFile);

        $writer->prepare();

        $writer->writeItem(['first', 'last']);

        $writer->finish();

        $reader = ReaderFactory::create($type);
        $reader->open($file);
        $spoutReader = new SpoutReader($reader);

        $this->assertCount(1, $spoutReader);
    }

    /**
     * @dataProvider writerTypeProvider
     */
    public function testHeaderNotPrependedByDefault($type)
    {
        $file = tempnam(sys_get_temp_dir(), null);
        $writerFile = WriterFactory::create($type);
        $writerFile->openToFile($file);
        $writer = new SpoutWriter($writerFile);
        $writer->prepare();
        $writer->writeItem([
            'col 1 name'=>'col 1 value',
            'col 2 name'=>'col 2 value',
            'col 3 name'=>'col 3 value'
        ]);
        $writer->finish();

        $reader = ReaderFactory::create($type);
        $reader->open($file);
        $spoutReader = new SpoutReader($reader);

        $rows = iterator_to_array($spoutReader);
        # Values should be at first line
        $this->assertEquals(['col 1 value', 'col 2 value', 'col 3 value'], $rows[1]);
    }

    /**
     * @dataProvider writerTypeProvider
     */
    public function testHeaderPrependedWhenOptionSetToTrue($type)
    {
        $file = tempnam(sys_get_temp_dir(), null);
        $writerFile = WriterFactory::create($type);
        $writerFile->openToFile($file);

        $writer = new SpoutWriter($writerFile, null, true);
        $writer->prepare();
        $writer->writeItem([
            'col 1 name'=>'col 1 value',
            'col 2 name'=>'col 2 value',
            'col 3 name'=>'col 3 value'
        ]);
        $writer->finish();

        $reader = ReaderFactory::create($type);
        $reader->open($file);
        $spoutReader = new SpoutReader($reader);
        $rows = iterator_to_array($spoutReader);

        # Check column names at first line
        $this->assertEquals(['col 1 name', 'col 2 name', 'col 3 name'], $rows[1]);

        # Check values at second line
        $this->assertEquals(['col 1 value', 'col 2 value', 'col 3 value'], $rows[2]);
    }

    /**
     * @param \Box\Spout\Reader\ReaderInterface $reader
     * @param string $sheetTitle
     *
     * @return \Box\Spout\Reader\SheetInterface
     */
    private function getSheetByTitle(ReaderInterface $reader, $sheetTitle)
    {
        /** @var \Box\Spout\Reader\XLSX\Sheet $sheet */
        foreach ($reader->getSheetIterator() as $sheet) {
            if ($sheet->getName() === $sheetTitle) {
                return $sheet;
            }
        }

        throw new \RuntimeException(sprintf('The sheet "%s" does not exist', $sheetTitle));
    }
}

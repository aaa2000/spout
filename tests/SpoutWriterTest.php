<?php

namespace Port\Spout\Tests;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Reader\XLSX\Reader as XlsxReader;
use Port\Spout\SpoutWriter;

class SpoutWriterTest extends \PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        if (!extension_loaded('zip')) {
            $this->markTestSkipped();
        }
    }

    public function testWriteItemAppendWithSheetTitle()
    {
        $file = tempnam(sys_get_temp_dir(), null);
        $writer = new SpoutWriter(new \SplFileObject($file, 'w'), 'Sheet 1');

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

        /** @var \Box\Spout\Reader\XLSX\Reader $excel */
        $excel = ReaderFactory::create(Type::XLSX);
        $excel->open($file);

        $sheetOne = $this->getSheetByTitle($excel, 'Sheet 1');
        $this->assertCount(3, $sheetOne->getRowIterator());
    }

    public function testWriteItemWithoutSheetTitle()
    {
        $outputFile = new \SplFileObject(tempnam(sys_get_temp_dir(), null));
        $writer = new SpoutWriter($outputFile);

        $writer->prepare();

        $writer->writeItem(['first', 'last']);

        $writer->finish();
    }

    public function testHeaderNotPrependedByDefault()
    {
        $file = tempnam(sys_get_temp_dir(), null);

        $writer = new SpoutWriter(new \SplFileObject($file, 'w'), null, Type::XLSX);
        $writer->prepare();
        $writer->writeItem([
            'col 1 name'=>'col 1 value',
            'col 2 name'=>'col 2 value',
            'col 3 name'=>'col 3 value'
        ]);
        $writer->finish();

        /** @var \Box\Spout\Reader\XLSX\Reader $excel */
        $excel = ReaderFactory::create(Type::XLSX);
        $excel->open($file);
        /** @var  \Box\Spout\Reader\XLSX\Sheet $sheet */
        $sheet = current(iterator_to_array($excel->getSheetIterator()));
        $rows = iterator_to_array($sheet->getRowIterator());

        # Values should be at first line
        $this->assertEquals(['col 1 value', 'col 2 value', 'col 3 value'], $rows[1]);
    }

    public function testHeaderPrependedWhenOptionSetToTrue()
    {
        $file = tempnam(sys_get_temp_dir(), null);

        $writer = new SpoutWriter(new \SplFileObject($file, 'w'), null, Type::XLSX, true);
        $writer->prepare();
        $writer->writeItem([
            'col 1 name'=>'col 1 value',
            'col 2 name'=>'col 2 value',
            'col 3 name'=>'col 3 value'
        ]);
        $writer->finish();

        /** @var \Box\Spout\Reader\XLSX\Reader $excel */
        $excel = ReaderFactory::create(Type::XLSX);
        $excel->open($file);
        /** @var  \Box\Spout\Reader\XLSX\Sheet $sheet */
        /** @var  \Box\Spout\Reader\XLSX\Sheet $sheet */
        $sheet = current(iterator_to_array($excel->getSheetIterator()));
        $rows = iterator_to_array($sheet->getRowIterator());

        # Check column names at first line
        $this->assertEquals(['col 1 name', 'col 2 name', 'col 3 name'], $rows[1]);

        # Check values at second line
        $this->assertEquals(['col 1 value', 'col 2 value', 'col 3 value'], $rows[2]);
    }

    /**
     * @param \Box\Spout\Reader\XLSX\Reader $excel
     * @param string $sheetTitle
     *
     * @return \Box\Spout\Reader\XLSX\Sheet
     */
    private function getSheetByTitle(XlsxReader $excel, $sheetTitle)
    {
        /** @var \Box\Spout\Reader\XLSX\Sheet $sheet */
        foreach ($excel->getSheetIterator() as $sheet) {
            if ($sheet->getName() === $sheetTitle) {
                return $sheet;
            }
        }

        throw new \RuntimeException(sprintf('The sheet "%s" does not exist', $sheetTitle));
    }
}

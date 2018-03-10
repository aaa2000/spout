<?php

namespace Port\Spout;

use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Box\Spout\Writer\WriterInterface;
use Port\Writer;

/**
 * Writes to an Excel file
 */
class SpoutWriter implements Writer
{
    /**
     * @var null|string
     */
    protected $sheet;
    /**
     * @var boolean
     */
    protected $prependHeaderRow;
    /**
     * @var \Box\Spout\Writer\WriterInterface|\Box\Spout\Writer\AbstractMultiSheetsWriter
     */
    protected $writer;
    /**
     * @var integer
     */
    protected $row = 1;

    /**
     * Warning: the file data will be overwritten
     * @see http://opensource.box.com/spout/guides/add-data-to-existing-spreadsheet/
     * @see http://opensource.box.com/spout/guides/edit-existing-spreadsheet/
     *
     * @param \Box\Spout\Writer\WriterInterface $writer
     * @param string $sheet Sheet title (optional)
     * @param boolean $prependHeaderRow
     */
    public function __construct(WriterInterface $writer, $sheet = null, $prependHeaderRow = false)
    {
        $this->writer = $writer;
        $this->sheet = $sheet;
        $this->prependHeaderRow = $prependHeaderRow;
    }

    /**
     * {@inheritdoc}
     */
    public function prepare()
    {
        if (null !== $this->sheet && $this->writer instanceof AbstractMultiSheetsWriter) {
            $this->writer->getCurrentSheet()->setName($this->sheet);
        }
    }

    /**
     * {@inheritdoc}
     */
    public function writeItem(array $item)
    {
        if ($this->prependHeaderRow && 1 === $this->row) {
            $headers = array_keys($item);
            $this->writer->addRow($headers);
            $this->row++;
        }
        $values = array_values($item);
        $this->writer->addRow($values);
        $this->row++;
    }

    /**
     * {@inheritdoc}
     */
    public function finish()
    {
        $this->writer->close();
    }
}

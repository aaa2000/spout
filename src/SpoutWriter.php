<?php

namespace Port\Spout;

use Box\Spout\Writer\AbstractMultiSheetsWriter;
use Box\Spout\Writer\WriterFactory;
use Port\Writer;
use Box\Spout\Common\Type;

/**
 * Writes to an Excel file
 */
class SpoutWriter implements Writer
{
    /**
     * @var string
     */
    protected $filename;
    /**
     * @var null|string
     */
    protected $sheet;
    /**
     * @var string
     */
    protected $type;
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
     * @param \SplFileObject $file  File
     * @param string         $sheet Sheet title (optional)
     * @param string         $type  Excel file type (defaults to XLSX)
     * @param boolean        $prependHeaderRow
     */
    public function __construct(\SplFileObject $file, $sheet = null, $type = Type::XLSX, $prependHeaderRow = false)
    {
        $this->filename = $file->getPathname();
        $this->sheet = $sheet;
        $this->type = $type;
        $this->prependHeaderRow = $prependHeaderRow;
    }

    /**
     * {@inheritdoc}
     */
    public function prepare()
    {
        /** @var \Box\Spout\Writer\WriterInterface|\Box\Spout\Writer\AbstractMultiSheetsWriter $writer */
        $this->writer = WriterFactory::create($this->type);
        $this->writer->openToFile($this->filename);

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

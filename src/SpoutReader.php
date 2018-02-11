<?php

namespace Port\Spout;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Port\Reader\CountableReader;
use SeekableIterator;


/**
 * Reads Excel files with the help of Spout
 *
 * Spout must be installed.
 *
 * @link http://opensource.box.com/spout/
 * @link https://github.com/box/spout
 */
class SpoutReader implements CountableReader, SeekableIterator
{
    /**
     * @var \Box\Spout\Reader\XLSX\Sheet
     */
    protected $sheet;

    /**
     * @var integer
     */
    protected $headerRowNumber;

    /**
     * @var array
     */
    protected $columnHeaders;

    /**
     * Total number of rows
     *
     * @var integer
     */
    protected $count;

    /**
     * @param \SplFileObject $file            Excel file
     * @param integer        $headerRowNumber Optional number of header row
     * @param integer        $activeSheet     Index of active sheet to read from
     */
    public function __construct(\SplFileObject $file, $headerRowNumber = null, $activeSheet = null)
    {
        $reader = ReaderFactory::create(Type::XLSX);
        $reader->open($file->getPathname());

        if (null !== $activeSheet) {
            foreach ($reader->getSheetIterator() as $sheet) {
                if ($sheet->getIndex() === $activeSheet) {
                    break;
                }
            }
        }

        $this->sheet = $reader->getSheetIterator()->current();

        if (null !== $headerRowNumber) {
            $this->setHeaderRowNumber($headerRowNumber);
        }
    }

    /**
     * Return the current row as an array
     *
     * If a header row has been set, an associative array will be returned
     *
     * @return array
     */
    public function current()
    {
        $row = $this->sheet->getRowIterator()->current();

        // If the CSV has column headers, use them to construct an associative
        // array for the columns in this line
        if (!empty($this->columnHeaders)) {
            // Count the number of elements in both: they must be equal.
            // If not, ignore the row
            if (count($this->columnHeaders) === count($row)) {
                return array_combine(array_values($this->columnHeaders), $row);
            }
        } else {
            // Else just return the column values
            return $row;
        }
    }

    /**
     * Get column headers
     *
     * @return array
     */
    public function getColumnHeaders()
    {
        return $this->columnHeaders;
    }

    /**
     * Set column headers
     *
     * @param array $columnHeaders
     */
    public function setColumnHeaders(array $columnHeaders)
    {
        $this->columnHeaders = $columnHeaders;
    }

    /**
     * Rewind the file pointer
     *
     * If a header row has been set, the pointer is set just below the header
     * row. That way, when you iterate over the rows, that header row is
     * skipped.
     */
    public function rewind()
    {
        $this->sheet->getRowIterator()->rewind();
        if (null !== $this->headerRowNumber) {
            $dataRowNumber = $this->headerRowNumber + 1;
            foreach ($this->sheet->getRowIterator() as $index => $row) {
                if ($index === $dataRowNumber) {
                    break;
                }
            }
        }
    }

    /**
     * Set header row number
     *
     * @param integer $rowNumber Number of the row that contains column header names
     */
    public function setHeaderRowNumber($rowNumber)
    {
        $this->headerRowNumber = $rowNumber;
        foreach ($this->sheet->getRowIterator() as $index => $row) {
            if ($index === $rowNumber) {
                $this->columnHeaders = $row;
                break;
            }
        }
        $this->columnHeaders = $this->worksheet[$rowNumber];
    }

    /**
     * {@inheritdoc}
     */
    public function next()
    {
        $this->sheet->getRowIterator()->next();
    }

    /**
     * {@inheritdoc}
     */
    public function valid()
    {
        return $this->sheet->getRowIterator()->valid();
    }

    /**
     * {@inheritdoc}
     */
    public function key()
    {
        return $this->sheet->getRowIterator()->key();
    }

    /**
     * {@inheritdoc}
     */
    public function seek($position)
    {
        foreach ($this->sheet->getRowIterator() as $index => $row) {
            if ($index === $position) {
                break;
            }
        }
    }

    /**
     * {@inheritdoc}
     */
    public function count()
    {
        $count = iterator_count($this->sheet);

        return $count - (int) $this->headerRowNumber;
    }

    /**
     * Get a row
     *
     * @param integer $number
     *
     * @return array
     */
    public function getRow($number)
    {
        $this->seek($number);

        return $this->current();
    }
}

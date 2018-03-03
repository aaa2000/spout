<?php

namespace Port\Spout;

use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use OutOfBoundsException;
use Port\Exception\ReaderException;
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
     * @param \SplFileObject $file Excel file
     * @param int $headerRowNumber Optional number of header row
     * @param int $activeSheet Index of active sheet to read from
     * @param bool $shouldPreserveEmptyRows Sets whether empty rows should be returned or skipped
     *
     * @throws ReaderException
     */
    public function __construct(\SplFileObject $file, $headerRowNumber = null, $activeSheet = null, $shouldPreserveEmptyRows = true)
    {
        $reader = $this->createReaderForFile($file, $shouldPreserveEmptyRows);

        $activeSheet = null === $activeSheet ? 0 : (int) $activeSheet;
        foreach ($reader->getSheetIterator() as $sheet) {
            if ($sheet->getIndex() === $activeSheet) {
                break;
            }
        }

        if (!$reader->getSheetIterator()->valid()) {
            throw new ReaderException(sprintf('Sheet at index %d is not found', $activeSheet));
        }
        $this->sheet = $reader->getSheetIterator()->current();

        if (null !== $headerRowNumber) {
            $this->setHeaderRowNumber($headerRowNumber);
        }
    }

    /**
     * @param \SplFileObject $file
     * @param $shouldPreserveEmptyRows
     *
     * @return \Box\Spout\Reader\XLSX\Reader
     */
    private function createReaderForFile(\SplFileObject $file, $shouldPreserveEmptyRows)
    {
        /** @var \Box\Spout\Reader\XLSX\Reader $reader */
        $reader = ReaderFactory::create(Type::XLSX);
        $reader->setShouldPreserveEmptyRows($shouldPreserveEmptyRows);
        $reader->open($file->getPathname());

        return $reader;
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
            $this->seekIndex($this->headerRowNumber + 2);
        }
    }

    /**
     * Set header row number
     *
     * @param integer $rowNumber Number of the row that contains column header names
     */
    public function setHeaderRowNumber($rowNumber)
    {
        $rowNumber = (int) $rowNumber;
        $this->seekIndex($rowNumber + 1);
        $this->columnHeaders = $this->sheet->getRowIterator()->current();
        $this->headerRowNumber = $rowNumber;
        $this->sheet->getRowIterator()->next();
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
        $positionIndex = $position + 1;
        if (null !== $this->headerRowNumber) {
            $positionIndex += $this->headerRowNumber + 1;
        }

        $this->seekIndex($positionIndex);
    }

    /**
     * Seeks to a row index (one-based)
     *
     * @param $index
     */
    private function seekIndex($index)
    {
        $rowIterator = $this->sheet->getRowIterator();
        foreach ($rowIterator as $rowIndex => $row) {
            if ($rowIndex === $index) {
                return;
            }
        }

        throw new OutOfBoundsException(sprintf('Row number %d is out of range', $index));
    }

    /**
     * {@inheritdoc}
     */
    public function count()
    {
        $count = iterator_count($this->sheet->getRowIterator());

        if (null === $this->headerRowNumber) {
            return $count;
        }

        return $count - ($this->headerRowNumber + 1);
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

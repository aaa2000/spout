<?php

namespace Port\Spout;

use Box\Spout\Reader\Exception\IteratorNotRewindableException;
use Box\Spout\Reader\ReaderInterface;
use OutOfBoundsException;
use Port\Exception\ReaderException;
use Port\Reader\CountableReader;
use SeekableIterator;

/**
 * Reads Excel, ODS or CSV files with the help of Spout
 *
 * Spout must be installed.
 *
 * @link http://opensource.box.com/spout/
 * @link https://github.com/box/spout
 */
class SpoutReader implements CountableReader, SeekableIterator
{
    /**
     * @var \Box\Spout\Reader\ReaderInterface
     */
    protected $reader;

    /**
     * @var int|null
     */
    protected $activeSheet;

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
     * @var null|integer
     */
    protected $count;

    /**
     * @param ReaderInterface $reader
     * @param int $headerRowNumber Optional number of header row
     * @param int $activeSheet Index of active sheet to read from
     *
     * @throws ReaderException
     */
    public function __construct(ReaderInterface $reader, $headerRowNumber = null, $activeSheet = null)
    {
        $this->reader = $reader;
        $this->headerRowNumber = $headerRowNumber;
        $this->activeSheet = $activeSheet;
        $this->reinitialize();
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
        $this->reinitialize();
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
     *
     * WARNING: Seeking position in non-ascending order has performance impacts because the spreadsheet will be read
     * again from the beginning.
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
        if ($this->sheet->getRowIterator()->key() > $index) {
            $this->reinitialize();
        }

        $rowIterator = $this->sheet->getRowIterator();
        while ($rowIterator->valid()) {
            if ($rowIterator->key() === $index) {
                return;
            }
            $rowIterator->next();
        }

        throw new OutOfBoundsException(sprintf('Row number %d is out of range', $index));
    }

    /**
     * {@inheritdoc}
     */
    public function count()
    {
        $this->reinitialize();
        $count = 0;
        while ($this->sheet->getRowIterator()->valid()) {
            $count++;
            $this->sheet->getRowIterator()->next();
        }

        return $count;
    }

    /**
     * ODS RowIterator can only be done once, as it is not possible to read an XML file backwards.
     * @see \Box\Spout\Reader\ODS\RowIterator::rewind()
     *
     * The sheet iterator is then rewinded in order to instantiate a new sheet and so a new row iterator.
     * Move the row iterator on the row after the header row or on the first line if there is no headers.
     */
    private function reinitialize()
    {
        $activeSheet = null === $this->activeSheet ? 0 : (int)$this->activeSheet;
        foreach ($this->reader->getSheetIterator() as $sheet) {
            if ($sheet->getIndex() === $activeSheet) {
                break;
            }
        }

        if (!$this->reader->getSheetIterator()->valid()) {
            throw new ReaderException(sprintf('Sheet at index %d is not found', $activeSheet));
        }
        $this->sheet = $this->reader->getSheetIterator()->current();

        try {
            // Require for the XLSX format in order to load data
            $this->sheet->getRowIterator()->rewind();
        } catch (IteratorNotRewindableException $e) {
            // Ods row iterator can be rewinded only once
        }

        if (null !== $this->headerRowNumber) {
            $this->setHeaderRowNumber($this->headerRowNumber);
        } else {
            $this->seekIndex(1);
        }
    }
}

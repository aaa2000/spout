<?php

namespace Port\Spout;

use Port\Reader\ReaderFactory;

/**
 * Factory that creates Spout Readers
 */
class SpoutReaderFactory implements ReaderFactory
{
    /**
     * @var integer
     */
    protected $headerRowNumber;

    /**
     * @var integer
     */
    protected $activeSheet;
    /**
     * @var bool
     */
    private $shouldPreserveEmptyRows;

    /**
     * @param integer $headerRowNumber
     * @param integer $activeSheet
     */
    public function __construct($headerRowNumber = null, $activeSheet = null, $shouldPreserveEmptyRows = true)
    {
        $this->headerRowNumber = $headerRowNumber;
        $this->activeSheet = $activeSheet;
        $this->shouldPreserveEmptyRows = $shouldPreserveEmptyRows;
    }

    /**
     * @param \SplFileObject $file
     *
     * @return \Port\Spout\SpoutReader
     */
    public function getReader(\SplFileObject $file)
    {
        return new SpoutReader($file, $this->headerRowNumber, $this->activeSheet, $this->shouldPreserveEmptyRows);
    }
}

<?php

namespace Port\Spout;

use Box\Spout\Common\Type;
use Box\Spout\Reader\AbstractReader;
use Box\Spout\Reader\ReaderFactory as BoxSpoutReaderFactory;
use Port\Exception\UnexpectedValueException;
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
     * @param bool $shouldPreserveEmptyRows Sets whether empty rows should be returned or skipped
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
        $fileReader = $this->createReaderForFile($file);
        if ($fileReader instanceof AbstractReader) {
            $fileReader->setShouldPreserveEmptyRows($this->shouldPreserveEmptyRows);
        } elseif (true === $this->shouldPreserveEmptyRows) {
            throw new UnexpectedValueException(
                'The reader is not compatible with the parameter $shouldPreserveEmptyRows'
            );
        }
        $fileReader->open($file->getPathname());

        return new SpoutReader($fileReader, $this->headerRowNumber, $this->activeSheet);
    }

    private function createReaderForFile(\SplFileObject $file)
    {
        $extension = $file->getExtension();
        $reflectionClass = new \ReflectionClass(Type::class);
        $readerType = $reflectionClass->getConstant(strtoupper($extension));

        if (!$readerType) {
            throw new UnexpectedValueException(sprintf(
                'The extension "%s" of file "%s" is not managed by the factory',
                $extension,
                $file->getPathname()
            ));
        }

        return BoxSpoutReaderFactory::create($readerType);
    }
}

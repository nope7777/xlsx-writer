<?php

namespace XLSXWriter\Writer;

use XLSXWriter\Workbook;

/**
 * Class AbstractFile
 */
abstract class AbstractFile
{

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * @var string
     */
    protected $filePath;

    /**
     * AbstractFile constructor.
     *
     * @param Workbook $workbook
     */
    public function __construct(Workbook $workbook)
    {
        $this->workbook = $workbook;
    }

    /**
     * @return string
     */
    abstract public function getContents(): ?string;

    /**
     * @param string $path
     */
    public function save(string $path): void
    {
        $filePath = $path . '/' . $this->filePath;

        $this->makeDirs($filePath);

        file_put_contents($filePath, $this->getContents());
    }

    /**
     * @param string $path
     */
    protected function makeDirs(string $path): void
    {
        $dir = pathinfo($path , PATHINFO_DIRNAME);

        if (!is_dir($dir)) {
            $this->makeDirs($dir);
            mkdir($dir);
        }
    }

}
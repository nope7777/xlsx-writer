<?php

namespace XLSXWriter;

use XLSXWriter\Writer\AbstractFile;
use XLSXWriter\Writer\File;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;
use ZipArchive;

/**
 * Class Writer
 */
class Writer
{

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * @var AbstractFile[]
     */
    protected $generators = [
        File\ContentTypes::class,
        File\DocProps\App::class,
        File\DocProps\Core::class,
        File\Rels\Rels::class,
        File\XL\Styles::class,
        File\XL\Workbook::class,
        File\XL\Rels\Workbook::class,
        File\XL\Theme\Theme::class,
        File\XL\Worksheet\Sheet::class,
    ];

    /**
     * Writer constructor.
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
    protected function getAvailableTempPath(): string
    {
        do {
            $path = sys_get_temp_dir() . '/' . uniqid('xlsx_');
        } while (is_dir($path));

        mkdir($path);

        return realpath($path);
    }

    /**
     * @param string $path
     */
    protected function saveRaw(string $path): void
    {
        foreach ($this->generators as $generatorClass) {
            /** @var AbstractFile $generator */
            $generator = new $generatorClass ($this->workbook);
            $generator->save($path);
        }
    }

    /**
     * @param string $dirPath
     * @param string $filePath
     * @throws \Exception
     */
    protected function zip(string $dirPath, string $filePath): void
    {
        $zip = new ZipArchive;
        if ($zip->open($filePath, \ZipArchive::CREATE) !== true) {
            throw new \Exception ("Cannot create zip archive (path: " . $filePath . ")");
        }

        /** @var \SplFileInfo[] $files */
        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($dirPath),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach ($files as $name => $file) {
            if (! $file->isDir()) {
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($dirPath) + 1);

                $zip->addFile($filePath, $relativePath);
            }
        }

        $zip->close();
    }

    /**
     * @param string $path
     */
    public function save(string $path): void
    {
        $tempPath = $this->getAvailableTempPath();
        $this->saveRaw($tempPath);
        $this->zip($tempPath, $path);
    }

    /**
     * @param string $name
     */
    public function output(string $name): void
    {
        $tmp_path = tempnam(sys_get_temp_dir(), '');

        $this->save($tmp_path);

        header('Content-Description: File Transfer');
        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="' . $name . '"');
        header('Expires: 0');
        header('Cache-Control: must-revalidate');
        header('Pragma: public');
        header('Content-Length: ' . filesize($tmp_path));
        readfile($tmp_path);

        unlink($tmp_path);
    }

}
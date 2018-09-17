<?php

namespace XLSXWriter;

use XLSXWriter\Workbook\Sheet;
use XLSXWriter\Workbook\Style;

/**
 * Class Workbook
 */
class Workbook
{

    /**
     * @var Sheet[]
     */
    protected $sheets = [];

    /**
     * @var Style[]
     */
    protected $styles = [];

    /**
     * @var Style
     */
    protected $defaultStyle;

    /**
     * @var int
     */
    protected $activeSheetIndex = 0;

    /**
     * Workbook constructor.
     */
    public function __construct()
    {
        $this->defaultStyle = new Style;
    }

    /**
     * @param string|null $title
     * @return Sheet
     */
    public function addSheet(string $title = null): Sheet
    {
        $sheet = new Sheet ($this, $title);
        $this->sheets[] = $sheet;

        return $sheet;
    }

    /**
     * @return Sheet[]
     */
    public function getSheets(): array
    {
        return $this->sheets;
    }

    /**
     * @return int
     */
    public function getSheetsCount(): int
    {
        return count($this->sheets);
    }

    /**
     * @return Style
     */
    public function getDefaultStyle(): Style
    {
        return $this->defaultStyle;
    }

    /**
     * @return Style
     */
    public function getNewStyle(): Style
    {
        return clone $this->defaultStyle;
    }

    /**
     * @param bool $withDefault
     * @return Style[]
     */
    public function getStyles(bool $withDefault = true): array
    {
        if ($withDefault) {
            return array_merge(
                [$this->defaultStyle],
                array_values($this->styles)
            );
        }

        return array_values($this->styles);
    }

    /**
     * Adds style to workbook and returns it index.
     *
     * @param Style $style
     * @return integer
     */
    public function addStyle(Style $style): int
    {
        // If style defined, return index
        $hash = $style->getHash();
        if (isset($this->styles[$hash])) {
            return array_search($hash, array_keys($this->styles)) + 1;
        }

        $this->styles[$hash] = clone $style;

        return count($this->styles);
    }

    /**
     * @return int
     */
    public function getActiveSheetIndex(): int
    {
        return $this->activeSheetIndex;
    }

    /**
     * @param int $index
     */
    public function setActiveSheetIndex(int $index): void
    {
        $this->activeSheetIndex = $index;
    }

    /**
     * @return Sheet
     */
    public function getActiveSheet(): Sheet
    {
        return $this->sheets[$this->activeSheetIndex];
    }

    /**
     * @param string $path
     */
    public function save(string $path): void
    {
        $writer = new Writer($this);
        $writer->save($path);
    }

    /**
     * @param string $name
     */
    public function output(string $name): void
    {
        $writer = new Writer($this);
        $writer->output($name);
    }

}
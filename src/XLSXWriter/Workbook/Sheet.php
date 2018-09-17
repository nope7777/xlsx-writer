<?php

namespace XLSXWriter\Workbook;

use XLSXWriter\Workbook;
use XLSXWriter\Workbook\Sheet\Row;

/**
 * Class Sheet
 */
class Sheet
{

    /**
     * @var Workbook
     */
    protected $workbook;

    /**
     * @var string
     */
    protected $title;

    /**
     * @var array
     */
    protected $columnsWidths = [];

    /**
     * @var array
     */
    protected $mergedCells = [];

    /**
     * @var int|null
     */
    protected $paneY;

    /**
     * @var string|null
     */
    protected $filters;

    /**
     * @var string
     */
    protected $selectionFrom = 'A1';

    /**
     * @var string
     */
    protected $selectionTo = 'A1';

    /**
     * @var int
     */
    protected $rowsCount = 0;

    /**
     * @var int
     */
    protected $maxColumnsPerRow = 0;

    /**
     * @var resource
     */
    protected $tempFilePointer;

    /**
     * Sheet constructor.
     *
     * @param Workbook $workbook
     * @param string|null $title
     */
    public function __construct(Workbook $workbook, string $title = null)
    {
        $this->workbook = $workbook;
        $this->title = $title ?? 'Sheet ' . ($this->workbook->getSheetsCount() + 1);

        $this->tempFilePointer = tmpfile();
    }

    /**
     * @return Workbook
     */
    public function getWorkbook(): Workbook
    {
        return $this->workbook;
    }

    /**
     * @param string $title
     */
    public function setTitle(string $title): void
    {
        $this->title = $title;
    }

    /**
     * @return string
     */
    public function getTitle(): string
    {
        return $this->title;
    }

    /**
     * @param int $column Starts with 1
     * @param float $width
     * @return void
     */
    public function setColumnWidth(int $column, float $width): void
    {
        $this->columnsWidths[$column] = $width;
    }

    /**
     * @return array
     */
    public function getColumnsWidths(): array
    {
        return $this->columnsWidths;
    }

    /**
     * @param string $from
     * @param string $to
     */
    public function mergeCells(string $from, string $to): void
    {
        $this->mergedCells[] = $from . ':' . $to;
    }

    /**
     * @return array
     */
    public function getMergedCells(): array
    {
        return $this->mergedCells;
    }

    /**
     * @return int|null
     */
    public function getPaneY(): ?int
    {
        return $this->paneY;
    }

    /**
     * @param int $paneY
     */
    public function setPaneY($paneY): void
    {
        $this->paneY = $paneY;
    }

    /**
     * @return string|null
     */
    public function getFilters(): ?string
    {
        return $this->filters;
    }

    /**
     * @param string $filters
     */
    public function setFilters(string $filters): void
    {
        $this->filters = $filters;
    }

    /**
     * @param string $from
     * @param string|null $to
     */
    public function setSelection(string $from, string $to = null): void
    {
        $this->selectionFrom = $from;
        $this->selectionTo = $to ?? $from;
    }

    /**
     * @return string
     */
    public function getSelectionFrom(): string
    {
        return $this->selectionFrom;
    }

    /**
     * @return string
     */
    public function getSelectionTo(): string
    {
        return $this->selectionTo;
    }

    /**
     * @param array|Row $row
     * @param \XLSXWriter\Workbook\Style|null $style
     */
    public function addRow($row, Style $style = null): void
    {
        if (! ($row instanceof Row)) {
            $row = new Row($row, $style);
        }

        if ($row->getCellsCount() > $this->maxColumnsPerRow) {
            $this->maxColumnsPerRow = $row->getCellsCount();
        }

        $this->rowsCount++;

        fwrite($this->tempFilePointer, $row->getXML($this->rowsCount, $this->workbook));
    }

    /**
     * @param array $rows
     * @param Style|null $style
     */
    public function addRows(array $rows, Style $style = null): void
    {
        foreach ($rows as $row) {
            $this->addRow($row, $style);
        }
    }

    /**
     * @param int $count
     */
    public function skipRows(int $count = 1): void
    {
        for ($i = 1; $i <= $count; $i++) {
            $this->addRow([]);
        }
    }

    /**
     * @return int
     */
    public function getRowsCount(): int
    {
        return $this->rowsCount;
    }

    /**
     * @return int
     */
    public function getCurrentRow(): int
    {
        return $this->rowsCount + 1;
    }

    /**
     * @return int
     */
    public function getMaxColumnsPerRow(): int
    {
        return $this->maxColumnsPerRow;
    }

    /**
     * @return resource
     */
    public function getTempFilePointer()
    {
        return $this->tempFilePointer;
    }

}
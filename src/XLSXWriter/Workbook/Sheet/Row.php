<?php

namespace XLSXWriter\Workbook\Sheet;

use XLSXWriter\Workbook;
use XLSXWriter\Workbook\Style;
use XLSXWriter\Workbook\Sheet\Row\Cell;
use XMLWriter;

/**
 * Class Row
 */
class Row
{

    /**
     * @var Cell[]
     */
    protected $cells = [];

    /**
     * Row constructor.
     *
     * @param array $cells
     * @param Style|null $style
     */
    public function __construct(array $cells, Style $style = null)
    {
        $this->addCells($cells, $style);
    }

    /**
     * @param mixed $value
     * @param Style|null $style
     */
    public function addCell($value, Style $style = null): void
    {
        $this->cells[] = $value instanceof Cell ? $value : new Cell($value, $style);
    }

    /**
     * @param mixed $values
     * @param Style|null $style
     */
    public function addCells($values, Style $style = null): void
    {
        foreach ($values as $value) {
            $this->addCell($value, $style);
        }
    }

    /**
     * @return Cell[]
     */
    public function getCells(): array
    {
        return $this->cells;
    }

    /**
     * @return int
     */
    public function getCellsCount(): int
    {
        return count($this->cells);
    }

    /**
     * @param int $rowIndex
     * @param Workbook $workbook
     * @return string
     */
    public function getXML(int $rowIndex, Workbook $workbook): string
    {
        if (empty($this->cells)) {
            return '';
        }

        $xw = new XMLWriter;
        $xw->openMemory();
        $xw->startElement('row');
        $xw->writeAttribute('r', $rowIndex);
        $xw->writeAttribute('spans', '1:' . $this->getCellsCount());
        foreach ($this->cells as $cellIndex => $cell) {
            $xw->writeRaw($cell->getXML($rowIndex, $cellIndex, $workbook));
        }
        $xw->endElement();

        return $xw->flush();
    }

}
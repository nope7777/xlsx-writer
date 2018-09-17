<?php

namespace XLSXWriter\Workbook\Sheet\Row;

use XLSXWriter\Workbook;
use XLSXWriter\Workbook\Style;
use XLSXWriter\Helper\ExcelHelper;
use XMLWriter;
use DateTime;

/**
 * Class Cell
 */
class Cell
{

    /**
     * @var mixed
     */
    protected $value;

    /**
     * @var Style|null
     */
    protected $style;

    /**
     * Cell constructor.
     *
     * @param mixed|null $value
     * @param Style|null $style
     */
    public function __construct($value = null, Style $style = null)
    {
        $this->value = $value;
        $this->style = $style;
    }

    /**
     * @return mixed|null
     */
    public function getValue()
    {
        return $this->value;
    }

    /**
     * @param int $rowIndex
     * @param int $cellIndex
     * @param Workbook $workbook
     * @return string
     */
    public function getXML(int $rowIndex, int $cellIndex, Workbook $workbook): string
    {
        $styleIndex = $this->style ? $workbook->addStyle($this->style) : null;

        $columnLetter = ExcelHelper::indexToLetter($cellIndex);

        $xw = new XMLWriter;
        $xw->openMemory();
        $xw->startElement('c');
        $xw->writeAttribute('r', $columnLetter . $rowIndex);
        if ($styleIndex !== null) {
            $xw->writeAttribute('s', $styleIndex);
        }

        if (is_string($this->value)) {
            $xw->writeAttribute('t', 'inlineStr');
            $xw->startElement('is');
                $xw->startElement('t');
                    $xw->writeRaw(ExcelHelper::escapeString($this->value));
                $xw->endElement();
            $xw->endElement();
        }
        elseif (is_bool($this->value)) {
            $xw->writeAttribute('t', 'b');
            $xw->writeElement('v', intval($this->value));
        }
        elseif (is_integer($this->value) || is_float($this->value)) {
            $xw->writeElement('v', $this->value);
        }
        elseif (is_object($this->value) && $this->value instanceof DateTime) {
            $xw->writeElement('v', ExcelHelper::PHPToExcel($this->value));
        }
        // If don't match any above condition and don't have style, then skip
        elseif ($this->style === null) {
            return '';
        }

        $xw->endElement();

        return $xw->flush();
    }

}
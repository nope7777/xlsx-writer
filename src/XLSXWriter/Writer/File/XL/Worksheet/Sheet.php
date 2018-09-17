<?php

namespace XLSXWriter\Writer\File\XL\Worksheet;

use XLSXWriter\Writer\AbstractFile;
use XLSXWriter\Workbook;
use XLSXWriter\Helper\ExcelHelper;
use XMLWriter;

/**
 * Class Sheet
 */
class Sheet extends AbstractFile
{

    /**
     * @return null|string
     */
    public function getContents(): ?string
    {
        return null;
    }

    /**
     * @param string $path
     */
    public function save(string $path): void
    {
        $this->makeDirs($path . '/xl/worksheets/sheet.xml');

        foreach ($this->workbook->getSheets() as $index => $sheet) {
            $filePath = $path . '/xl/worksheets/sheet' . ($index + 1) . '.xml';

            $pointer = fopen($filePath, 'w');



            $xw = new XMLWriter;
            $xw->openMemory();
            $xw->startDocument('1.0', 'UTF-8', 'yes');
            $xw->startElement('worksheet');
            $xw->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
            $xw->writeAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
            $xw->writeAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
            $xw->writeAttribute('mc:Ignorable', 'x14ac');
            $xw->writeAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');

            $xw->startElement('dimension');
            $xw->writeAttribute('ref', $this->getSheetDimension($sheet));
            $xw->endElement();

            $xw->startElement('sheetViews');
            $xw->startElement('sheetView');
            if ($index == $this->workbook->getActiveSheetIndex()) {
                $xw->writeAttribute('tabSelected', '1');
            }
            $xw->writeAttribute('workbookViewId', '0');

            // pane
            if ($sheet->getPaneY()) {
                $xw->startElement('pane');
                $xw->writeAttribute('ySplit', $sheet->getPaneY());
                $xw->writeAttribute('topLeftCell', 'A' . ($sheet->getPaneY() + 1));
                $xw->writeAttribute('activePane', 'bottomLeft');
                $xw->writeAttribute('state', 'frozen');
                $xw->endElement();
            }

            // selection
            $xw->startElement('selection');
            $xw->writeAttribute('activeCell', $sheet->getSelectionFrom());
            $xw->writeAttribute('sqref', $sheet->getSelectionTo());
            $xw->endElement();

            $xw->endElement();
            $xw->endElement();

            $xw->startElement('sheetFormatPr');
            $xw->writeAttribute('defaultRowHeight', '15');
            $xw->writeAttribute('x14ac:dyDescent', '0.25');
            $xw->endElement();

            // columns widths
            if ($sheet->getColumnsWidths()) {
                $xw->startElement('cols');
                foreach ($sheet->getColumnsWidths() as $column_index => $column_width) {
                    $xw->startElement('col');
                    $xw->writeAttribute('min', $column_index);
                    $xw->writeAttribute('max', $column_index);
                    $xw->writeAttribute('width', $column_width);
                    $xw->writeAttribute('customWidth', '1');
                    $xw->endElement();
                }
                $xw->endElement();
            }

            // Sheet data
            $xw->startElement('sheetData');

            if ($sheet->getRowsCount()) {
                // to close sheetData tag
                $xw->writeRaw('');

                // Write XMLWriter data to file and clear buffer
                fwrite($pointer, $xw->flush(true));

                // Write sheet temp data to sheet file
                rewind($sheet->getTempFilePointer());
                while (!feof($sheet->getTempFilePointer())) {
                    fwrite($pointer, fread($sheet->getTempFilePointer(), 10 * 1024 * 1024));
                }
            }

            $xw->endElement();

            // Auto filter
            if ($sheet->getFilters()) {
                $xw->startElement('autoFilter');
                $xw->writeAttribute('ref', $sheet->getFilters());
                $xw->endElement();
            }

            // Merged cells
            if (!empty($sheet->getMergedCells())) {
                $xw->startElement('mergeCells');
                $xw->writeAttribute('count', count($sheet->getMergedCells()));

                foreach ($sheet->getMergedCells() as $merged_cells) {
                    $xw->startElement('mergeCell');
                    $xw->writeAttribute('ref', $merged_cells);
                    $xw->endElement();
                }

                $xw->endElement();
            }

            $xw->startElement('pageMargins');
            $xw->writeAttribute('left', '0.7');
            $xw->writeAttribute('right', '0.7');
            $xw->writeAttribute('top', '0.75');
            $xw->writeAttribute('bottom', '0.75');
            $xw->writeAttribute('header', '0.3');
            $xw->writeAttribute('footer', '0.3');
            $xw->endElement();

            $xw->endElement();

            fwrite($pointer, $xw->flush(true));
            fclose($pointer);
        }
    }

    /**
     * @param Workbook\Sheet $sheet
     * @return string
     */
    protected function getSheetDimension(Workbook\Sheet $sheet): string
    {
        $dimension = 'A1';
        if ($sheet->getRowsCount() > 1 || $sheet->getMaxColumnsPerRow() > 1) {
            $dimension .= ':' . ExcelHelper::indexToLetter($sheet->getMaxColumnsPerRow() - 1) . $sheet->getRowsCount();
        }

        return $dimension;
    }

}
<?php

namespace XLSXWriter\Writer\File\XL;

use XLSXWriter\Writer\AbstractFile;
use DOMDocument;

/**
 * Class Workbook
 */
class Workbook extends AbstractFile
{

    /**
     * @var string
     */
    protected $filePath = 'xl/workbook.xml';

    /**
     * @return string
     */
    public function getContents(): string
    {
        $xml = new DOMDocument('1.0', 'UTF-8');
        $xml->xmlStandalone = true;

        $xmlWorkbook = $xml->createElement('workbook');
        $xmlWorkbook->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $xmlWorkbook->setAttribute('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
        $xmlWorkbook->setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        $xmlWorkbook->setAttribute('mc:Ignorable', 'x15');
        $xmlWorkbook->setAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');
        $xml->appendChild($xmlWorkbook);

        $xmlFileVersion = $xml->createElement('fileVersion');
        $xmlFileVersion->setAttribute('appName', 'xl');
        $xmlFileVersion->setAttribute('lastEdited', '7');
        $xmlFileVersion->setAttribute('lowestEdited', '7');
        $xmlFileVersion->setAttribute('rupBuild', '19001');
        $xmlWorkbook->appendChild($xmlFileVersion);

        $xmlBookViews = $xml->createElement('bookViews');
        $xmlWorkbook->appendChild($xmlBookViews);

        $xmlWorkbookView = $xml->createElement('workbookView');
        $xmlWorkbookView->setAttribute('xWindow', '0');
        $xmlWorkbookView->setAttribute('yWindow', '0');
        $xmlWorkbookView->setAttribute('windowWidth', '28770');
        $xmlWorkbookView->setAttribute('windowHeight', '11865');
        $xmlWorkbookView->setAttribute('activeTab', $this->workbook->getActiveSheetIndex());
        $xmlBookViews->appendChild($xmlWorkbookView);

        $xmlSheets = $xml->createElement('sheets');
        $xmlWorkbook->appendChild($xmlSheets);

        foreach ($this->workbook->getSheets() as $index => $sheet) {
            $xmlSheet = $xml->createElement('sheet');
            $xmlSheet->setAttribute('name', $sheet->getTitle());
            $xmlSheet->setAttribute('sheetId', $index + 1);
            $xmlSheet->setAttribute('r:id', 'rId' . ($index + 1));
            $xmlSheets->appendChild($xmlSheet);
        }

        $xmlCalcPr = $xml->createElement('calcPr');
        $xmlCalcPr->setAttribute('calcId', '162913');
        $xmlWorkbook->appendChild($xmlCalcPr);

        $xmlExtLst = $xml->createElement('extLst');
        $xmlWorkbook->appendChild($xmlExtLst);

        $xmlExl = $xml->createElement('ext');
        $xmlExl->setAttribute('uri', '{140A7094-0E35-4892-8432-C4D2E57EDEB5}');
        $xmlExl->setAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');
        $xmlExtLst->appendChild($xmlExl);

        $xmlX15WorkbookPr = $xml->createElement('x15:workbookPr');
        $xmlX15WorkbookPr->setAttribute('chartTrackingRefBase', '1');
        $xmlExl->appendChild($xmlX15WorkbookPr);

        return $xml->saveXML();
    }

}
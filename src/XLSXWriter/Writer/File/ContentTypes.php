<?php

namespace XLSXWriter\Writer\File;

use XLSXWriter\Writer\AbstractFile;
use DOMDocument;

/**
 * Class ContentTypes
 */
class ContentTypes extends AbstractFile
{

    /**
     * @var string
     */
    protected $filePath = '[Content_Types].xml';

    /**
     * @return string
     */
    public function getContents(): string
    {
        $xml = new DOMDocument('1.0', 'UTF-8');
        $xml->xmlStandalone = true;

        $xmlTypes = $xml->createElement('Types');
        $xmlTypes->setAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/content-types');
        $xml->appendChild($xmlTypes);

        $xmlDefault = $xml->createElement('Default');
        $xmlDefault->setAttribute('Extension', 'rels');
        $xmlDefault->setAttribute('ContentType', 'application/vnd.openxmlformats-package.relationships+xml');
        $xmlTypes->appendChild($xmlDefault);

        $xmlSecondDefault = $xml->createElement('Default');
        $xmlSecondDefault->setAttribute('Extension', 'xml');
        $xmlSecondDefault->setAttribute('ContentType', 'application/xml');
        $xmlTypes->appendChild($xmlSecondDefault);

        $xmlWorkbook = $xml->createElement('Override');
        $xmlWorkbook->setAttribute('PartName', '/xl/workbook.xml');
        $xmlWorkbook->setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml');
        $xmlTypes->appendChild($xmlWorkbook);

        foreach ($this->workbook->getSheets() as $index => $sheet) {
            $xmlSheet = $xml->createElement('Override');
            $xmlSheet->setAttribute('PartName', '/xl/worksheets/sheet' . ($index + 1) . '.xml');
            $xmlSheet->setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml');
            $xmlTypes->appendChild($xmlSheet);
        }

        $xmlTheme = $xml->createElement('Override');
        $xmlTheme->setAttribute('PartName', '/xl/theme/theme1.xml');
        $xmlTheme->setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.theme+xml');
        $xmlTypes->appendChild($xmlTheme);

        $xmlStyles = $xml->createElement('Override');
        $xmlStyles->setAttribute('PartName', '/xl/styles.xml');
        $xmlStyles->setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml');
        $xmlTypes->appendChild($xmlStyles);

        $xmlCore = $xml->createElement('Override');
        $xmlCore->setAttribute('PartName', '/docProps/core.xml');
        $xmlCore->setAttribute('ContentType', 'application/vnd.openxmlformats-package.core-properties+xml');
        $xmlTypes->appendChild($xmlCore);

        $xmlApp = $xml->createElement('Override');
        $xmlApp->setAttribute('PartName', '/docProps/app.xml');
        $xmlApp->setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.extended-properties+xml');
        $xmlTypes->appendChild($xmlApp);

        return $xml->saveXML();
    }

}
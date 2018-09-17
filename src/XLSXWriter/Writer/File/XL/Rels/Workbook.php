<?php

namespace XLSXWriter\Writer\File\XL\Rels;

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
    protected $filePath = 'xl/_rels/workbook.xml.rels';

    /**
     * @return string
     */
    public function getContents(): string
    {
        $xml = new DOMDocument('1.0', 'UTF-8');
        $xml->xmlStandalone = true;

        $xmlRelationships = $xml->createElement('Relationships');
        $xmlRelationships->setAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $xml->appendChild($xmlRelationships);

        $rid = 1;
        foreach ($this->workbook->getSheets() as $sheet) {
            $xmlRelationship = $xml->createElement('Relationship');
            $xmlRelationship->setAttribute('Id', 'rId' . $rid);
            $xmlRelationship->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet');
            $xmlRelationship->setAttribute('Target', 'worksheets/sheet' . $rid . '.xml');
            $xmlRelationships->appendChild($xmlRelationship);

            $rid++;
        }

        $xmlRelationship = $xml->createElement('Relationship');
        $xmlRelationship->setAttribute('Id', 'rId' . $rid++);
        $xmlRelationship->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles');
        $xmlRelationship->setAttribute('Target', 'styles.xml');
        $xmlRelationships->appendChild($xmlRelationship);

        $xmlRelationship = $xml->createElement('Relationship');
        $xmlRelationship->setAttribute('Id', 'rId' . $rid);
        $xmlRelationship->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme');
        $xmlRelationship->setAttribute('Target', 'theme/theme1.xml');
        $xmlRelationships->appendChild($xmlRelationship);

        return $xml->saveXML();
    }

}
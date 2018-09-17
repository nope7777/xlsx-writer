<?php

namespace XLSXWriter\Writer\File\Rels;

use XLSXWriter\Writer\AbstractFile;
use DOMDocument;

/**
 * Class Rels
 */
class Rels extends AbstractFile
{

    /**
     * @var string
     */
    protected $filePath = '_rels/.rels';

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

        $xmlRelationship3 = $xml->createElement('Relationship');
        $xmlRelationship3->setAttribute('Id', 'rId3');
        $xmlRelationship3->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties');
        $xmlRelationship3->setAttribute('Target', 'docProps/app.xml');
        $xmlRelationships->appendChild($xmlRelationship3);

        $xmlRelationship2 = $xml->createElement('Relationship');
        $xmlRelationship2->setAttribute('Id', 'rId2');
        $xmlRelationship2->setAttribute('Type', 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties');
        $xmlRelationship2->setAttribute('Target', 'docProps/core.xml');
        $xmlRelationships->appendChild($xmlRelationship2);

        $xmlRelationship1 = $xml->createElement('Relationship');
        $xmlRelationship1->setAttribute('Id', 'rId1');
        $xmlRelationship1->setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument');
        $xmlRelationship1->setAttribute('Target', 'xl/workbook.xml');
        $xmlRelationships->appendChild($xmlRelationship1);

        return $xml->saveXML();
    }

}
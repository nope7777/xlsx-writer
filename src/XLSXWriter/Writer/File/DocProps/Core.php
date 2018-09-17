<?php

namespace XLSXWriter\Writer\File\DocProps;

use XLSXWriter\Writer\AbstractFile;
use DOMDocument;

/**
 * Class Core
 */
class Core extends AbstractFile
{

    /**
     * @var string
     */
    protected $filePath = 'docProps/core.xml';

    /**
     * @return string
     */
    public function getContents(): string
    {
        $xml = new DOMDocument('1.0', 'UTF-8');
        $xml->xmlStandalone = true;

        $xmlCpCoreProperties = $xml->createElement('cp:coreProperties');
        $xmlCpCoreProperties->setAttribute('xmlns:cp', 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
        $xmlCpCoreProperties->setAttribute('xmlns:dc', 'http://purl.org/dc/elements/1.1/');
        $xmlCpCoreProperties->setAttribute('xmlns:dcterms', 'http://purl.org/dc/terms/');
        $xmlCpCoreProperties->setAttribute('xmlns:dcmitype', 'http://purl.org/dc/dcmitype/');
        $xmlCpCoreProperties->setAttribute('xmlns:xsi', 'http://www.w3.org/2001/XMLSchema-instance');
        $xml->appendChild($xmlCpCoreProperties);

        $xmlDcCreator = $xml->createElement('dc:creator');
        $xmlDcCreator->textContent = 'nope7777/xlsx-writer';
        $xmlCpCoreProperties->appendChild($xmlDcCreator);

        $xmlCpLastModifiedBy = $xml->createElement('cp:lastModifiedBy');
        //$xmlCpLastModifiedBy->textContent = 'nope7777/xlsx-writer';
        $xmlCpCoreProperties->appendChild($xmlCpLastModifiedBy);

        $xmlDctermsCreated = $xml->createElement('dcterms:created');
        $xmlDctermsCreated->setAttribute('xsi:type', 'dcterms:W3CDTF');
        $xmlDctermsCreated->textContent = date('Y-m-d') . 'T' . date('H:i:s') . 'Z';
        $xmlCpCoreProperties->appendChild($xmlDctermsCreated);

        $xmlDctermsModified = $xml->createElement('dcterms:modified');
        $xmlDctermsModified->setAttribute('xsi:type', 'dcterms:W3CDTF');
        $xmlDctermsModified->textContent = date('Y-m-d') . 'T' . date('H:i:s') . 'Z';
        $xmlCpCoreProperties->appendChild($xmlDctermsModified);

        return $xml->saveXML();
    }

}
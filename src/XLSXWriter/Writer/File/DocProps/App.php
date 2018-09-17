<?php

namespace XLSXWriter\Writer\File\DocProps;

use XLSXWriter\Writer\AbstractFile;
use DOMDocument;

/**
 * Class App
 */
class App extends AbstractFile
{

    /**
     * @var string
     */
    protected $filePath = 'docProps/app.xml';

    /**
     * @return string
     */
    public function getContents(): string
    {
        $xml = new DOMDocument('1.0', 'UTF-8');
        $xml->xmlStandalone = true;

        $xmlProperties = $xml->createElement('Properties');
        $xmlProperties->setAttribute('xmlns', 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties');
        $xmlProperties->setAttribute('xmlns:vt', 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes');
        $xml->appendChild($xmlProperties);

        $xmlApplication = $xml->createElement('Application');
        $xmlApplication->textContent = 'Microsoft Excel';
        $xmlProperties->appendChild($xmlApplication);

        $xmlDocSecurity = $xml->createElement('DocSecurity');
        $xmlDocSecurity->textContent = 0;
        $xmlProperties->appendChild($xmlDocSecurity);

        $xmlScaleCrop = $xml->createElement('ScaleCrop');
        $xmlScaleCrop->textContent = 'false';
        $xmlProperties->appendChild($xmlScaleCrop);

        $xmlHeadingPairs = $xml->createElement('HeadingPairs');
        $xmlProperties->appendChild($xmlHeadingPairs);

        $xmlVtVector = $xml->createElement('vt:vector');
        $xmlVtVector->setAttribute('size', '2');
        $xmlVtVector->setAttribute('baseType', 'variant');
        $xmlHeadingPairs->appendChild($xmlVtVector);

        $xmlVtVariant = $xml->createElement('vt:variant');
        $xmlVtVector->appendChild($xmlVtVariant);

        $xmlVtLpstr = $xml->createElement('vt:lpstr');
        $xmlVtLpstr->textContent = 'Sheets';
        $xmlVtVariant->appendChild($xmlVtLpstr);

        $xmlSecondVtVariant = $xml->createElement('vt:variant');
        $xmlVtVector->appendChild($xmlSecondVtVariant);

        $xmlVtI4 = $xml->createElement('vt:i4');
        $xmlVtI4->textContent = $this->workbook->getSheetsCount();
        $xmlSecondVtVariant->appendChild($xmlVtI4);

        $xmlTitlesOfParts = $xml->createElement('TitlesOfParts');
        $xmlProperties->appendChild($xmlTitlesOfParts);

        $xmlVtVector = $xml->createElement('vt:vector');
        $xmlVtVector->setAttribute('size', $this->workbook->getSheetsCount());
        $xmlVtVector->setAttribute('baseType', 'lpstr');
        $xmlTitlesOfParts->appendChild($xmlVtVector);

        foreach ($this->workbook->getSheets() as $index => $sheet) {
            $xmlVtLpstr = $xml->createElement('vt:lpstr');
            $xmlVtLpstr->textContent = $sheet->getTitle();
            $xmlVtVector->appendChild($xmlVtLpstr);
        }

        $xmlCompany = $xml->createElement('Company');
        //$xmlCompany->textContent = '';
        $xmlProperties->appendChild($xmlCompany);

        $xmlLinksUpToDate = $xml->createElement('LinksUpToDate');
        $xmlLinksUpToDate->textContent = 'false';
        $xmlProperties->appendChild($xmlLinksUpToDate);

        $xmlSharedDoc = $xml->createElement('SharedDoc');
        $xmlSharedDoc->textContent = 'false';
        $xmlProperties->appendChild($xmlSharedDoc);

        $xmlHyperlinksChanged = $xml->createElement('HyperlinksChanged');
        $xmlHyperlinksChanged->textContent = 'false';
        $xmlProperties->appendChild($xmlHyperlinksChanged);

        $xmlAppVersion = $xml->createElement('AppVersion');
        $xmlAppVersion->textContent = '16.0300';
        $xmlProperties->appendChild($xmlAppVersion);

        return $xml->saveXML();
    }

}
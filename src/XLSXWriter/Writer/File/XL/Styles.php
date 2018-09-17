<?php

namespace XLSXWriter\Writer\File\XL;

use XLSXWriter\Writer\AbstractFile;
use XLSXWriter\Workbook;
use XLSXWriter\Workbook\Style;

class Styles extends AbstractFile
{
    /**
     * @var string
     */
    protected $filePath = 'xl/styles.xml';

    /**
     * @var Style[]
     */
    protected $fonts = [];

    /**
     * @var Style[]
     */
    protected $fills = [];

    /**
     * @var Style[]
     */
    protected $borders = [];

    /**
     * Styles constructor.
     *
     * @param Workbook $workbook
     */
    public function __construct(Workbook $workbook)
    {
        parent::__construct($workbook);

        // First <fill> must be "none"
        $style = (new Style)->setFillPatternType('none');
        $this->fills[$style->getFillHash()] = $style;

        // Second <fill> must be "gray125"
        $style = (new Style)->setFillPatternType('gray125');
        $this->fills[$style->getFillHash()] = $style;

        // First <border> must be empty
        $style = new Style;
        $this->borders[$style->getBorderHash()] = $style;

        // Generating maps
        foreach ($this->workbook->getStyles() as $style) {
            $this->fonts[$style->getFontHash()] = $style;
            $this->fills[$style->getFillHash()] = $style;
            $this->borders[$style->getBorderHash()] = $style;
        }
    }

    /**
     * @return string
     */
    public function getContents(): string
    {
        $xml = new \DOMDocument('1.0', 'UTF-8');
        $xml->xmlStandalone = true;

        $xmlStyle = $xml->createElement('styleSheet');
        $xmlStyle->setAttribute('xmlns', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main');
        $xmlStyle->setAttribute('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006');
        $xmlStyle->setAttribute('mc:Ignorable', 'x14ac x16r2');
        $xmlStyle->setAttribute('xmlns:x14ac', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac');
        $xmlStyle->setAttribute('xmlns:x16r2', 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main');
        $xml->appendChild($xmlStyle);

        /**
         * Fonts
         */
        $xmlFonts = $xml->createElement('fonts');
        $xmlFonts->setAttribute('count', count($this->fonts));
        $xmlFonts->setAttribute('x14ac:knownFonts', '1');
        $xmlStyle->appendChild($xmlFonts);

        foreach ($this->fonts as $font) {
            $xmlFont = $xml->createElement('font');
            $xmlFonts->appendChild($xmlFont);

            $xmlFontSize = $xml->createElement('sz');
            $xmlFontSize->setAttribute('val', $font->getFontSize());
            $xmlFont->appendChild($xmlFontSize);

            $xmlFontName = $xml->createElement('name');
            $xmlFontName->setAttribute('val', $font->getFontName());
            $xmlFont->appendChild($xmlFontName);

            if ($font->getFontColor()) {
                $xmlFontColor = $xml->createElement('color');
                $xmlFontColor->setAttribute('rgb', 'FF' . $font->getFontColor());
                $xmlFont->appendChild($xmlFontColor);
            }

            if ($font->isFontBold()) {
                $xmlFontBold = $xml->createElement('b');
                $xmlFont->appendChild($xmlFontBold);
            }

            if ($font->isFontItalic()) {
                $xmlFontItalic = $xml->createElement('i');
                $xmlFont->appendChild($xmlFontItalic);
            }

            if ($font->isFontUnderline()) {
                $xmlFontUnderline = $xml->createElement('u');
                $xmlFont->appendChild($xmlFontUnderline);
            }

            if ($font->isFontStrikethrough()) {
                $xmlFontStrike = $xml->createElement('strike');
                $xmlFont->appendChild($xmlFontStrike);
            }
        }

        /**
         * Fills
         */
        $xmlFills = $xml->createElement('fills');
        $xmlFills->setAttribute('count', count($this->fills));
        $xmlStyle->appendChild($xmlFills);

        foreach ($this->fills as $fill) {
            $xmlFill = $xml->createElement('fill');
            $xmlFills->appendChild($xmlFill);

            $xmlFillPattern = $xml->createElement('patternFill');
            $xmlFillPattern->setAttribute('patternType', $fill->getFillPatternType());
            $xmlFill->appendChild($xmlFillPattern);

            if (!is_null($fill->getFillColor())) {
                $xmlFillPatternFgcolor = $xml->createElement('fgColor');
                $xmlFillPatternFgcolor->setAttribute('rgb', $fill->getFillColor());
                $xmlFillPattern->appendChild($xmlFillPatternFgcolor);
            }
        }

        /**
         * Borders
         */
        $xmlBorders = $xml->createElement('borders');
        $xmlBorders->setAttribute('count', count($this->borders));
        $xmlStyle->appendChild($xmlBorders);

        foreach ($this->borders as $border) {
            $xmlBorder = $xml->createElement('border');
            $xmlBorders->appendChild($xmlBorder);

            // left
            $xmlBorderLeft = $xml->createElement('left');
            if ($border->getBorderLeftStyle()) {
                $xmlBorderLeft->setAttribute('style', $border->getBorderLeftStyle());

                if ($border->getBorderLeftColor()) {
                    $xmlBorderLeftColor = $xml->createElement('color');
                    $xmlBorderLeftColor->setAttribute('rgb', $border->getBorderLeftColor());
                    $xmlBorderLeft->appendChild($xmlBorderLeftColor);
                }
            }
            $xmlBorder->appendChild($xmlBorderLeft);

            // right
            $xmlBorderRight = $xml->createElement('right');
            if ($border->getBorderRightStyle()) {
                $xmlBorderRight->setAttribute('style', $border->getBorderRightStyle());

                if ($border->getBorderRightColor()) {
                    $xmlBorderRightColor = $xml->createElement('color');
                    $xmlBorderRightColor->setAttribute('rgb', $border->getBorderRightColor());
                    $xmlBorderRight->appendChild($xmlBorderRightColor);
                }
            }
            $xmlBorder->appendChild($xmlBorderRight);

            // top
            $xmlBorderTop = $xml->createElement('top');
            if ($border->getBorderTopStyle()) {
                $xmlBorderTop->setAttribute('style', $border->getBorderTopStyle());

                if ($border->getBorderTopColor()) {
                    $xmlBorderTopColor = $xml->createElement('color');
                    $xmlBorderTopColor->setAttribute('rgb', $border->getBorderTopColor());
                    $xmlBorderTop->appendChild($xmlBorderTopColor);
                }
            }
            $xmlBorder->appendChild($xmlBorderTop);

            // bottom
            $xmlBorderBottom = $xml->createElement('bottom');
            if ($border->getBorderBottomStyle()) {
                $xmlBorderBottom->setAttribute('style', $border->getBorderBottomStyle());

                if ($border->getBorderBottomColor()) {
                    $xmlBorderBottomColor = $xml->createElement('color');
                    $xmlBorderBottomColor->setAttribute('rgb', $border->getBorderBottomColor());
                    $xmlBorderBottom->appendChild($xmlBorderBottomColor);
                }
            }
            $xmlBorder->appendChild($xmlBorderBottom);

            // diagonal
            $xmlBorderDiagonal = $xml->createElement('diagonal');
            if ($border->getBorderDiagonalStyle()) {
                $xmlBorder->setAttribute('diagonalUp', '1');
                $xmlBorder->setAttribute('diagonalDown', '1');

                $xmlBorderDiagonal->setAttribute('style', $border->getBorderDiagonalStyle());

                if ($border->getBorderDiagonalColor()) {
                    $xmlBorderDiagonalColor = $xml->createElement('color');
                    $xmlBorderDiagonalColor->setAttribute('rgb', $border->getBorderDiagonalColor());
                    $xmlBorderDiagonal->appendChild($xmlBorderDiagonalColor);
                }
            }
            $xmlBorder->appendChild($xmlBorderDiagonal);
        }

        /**
         * cellStyleXfs
         */
        $xmlStyleXfs = $xml->createElement('cellStyleXfs');
        $xmlStyleXfs->setAttribute('count', '1');
        $xmlStyle->appendChild($xmlStyleXfs);

        $xmlStyleXfsXf = $xml->createElement('xf');
        $xmlStyleXfsXf->setAttribute('numFmtId', '0');
        $xmlStyleXfsXf->setAttribute('fontId', '0');
        $xmlStyleXfsXf->setAttribute('fillId', '0');
        $xmlStyleXfsXf->setAttribute('borderId', '0');
        $xmlStyleXfs->appendChild($xmlStyleXfsXf);

        /**
         * cellXfs (Style)
         */
        $xmlCellXfs = $xml->createElement('cellXfs');
        $xmlCellXfs->setAttribute('count', count($this->workbook->getStyles()));
        $xmlStyle->appendChild($xmlCellXfs);

        foreach ($this->workbook->getStyles() as $style) {
            $xmlCellXfsXf = $xml->createElement('xf');
            $xmlCellXfsXf->setAttribute('numFmtId', '0');
            $xmlCellXfsXf->setAttribute('fontId', '0');
            $xmlCellXfsXf->setAttribute('fillId', '0');
            $xmlCellXfsXf->setAttribute('borderId', '0');
            $xmlCellXfsXf->setAttribute('xfId', '0');

            if ($style->getNumberFormat()) {
                $xmlCellXfsXf->setAttribute('numFmtId', $style->getNumberFormat());
                $xmlCellXfsXf->setAttribute('applyNumberFormat', '1');
            }

            //if ($style->fontDefined()) {
                $xmlCellXfsXf->setAttribute('fontId', $this->getFontId($style));
                $xmlCellXfsXf->setAttribute('applyFont', '1');
            //}

            //if ($style->fillDefined()) {
                $xmlCellXfsXf->setAttribute('fillId', $this->getFillId($style));
                $xmlCellXfsXf->setAttribute('applyFill', '1');
            //}

            //if ($style->bordersDefined()) {
                $xmlCellXfsXf->setAttribute('borderId', $this->getBorderId($style));
                $xmlCellXfsXf->setAttribute('applyBorder', '1');
            //}

            $xmlCellXfs->appendChild($xmlCellXfsXf);

            // Alignment
            if ($style->getHorizontalAlignment() || $style->getVerticalAlignment()) {
                $xmlCellXfsAlignment = $xml->createElement('alignment');
                if ($style->getHorizontalAlignment()) {
                    $xmlCellXfsAlignment->setAttribute('horizontal', $style->getHorizontalAlignment());
                }
                if ($style->getVerticalAlignment()) {
                    $xmlCellXfsAlignment->setAttribute('vertical', $style->getVerticalAlignment());
                }
                $xmlCellXfsXf->appendChild($xmlCellXfsAlignment);
            }
        }

        /**
         * cellStyles
         */
        $xmlCellStyles = $xml->createElement('cellStyles');
        $xmlCellStyles->setAttribute('count', '1');
        $xmlStyle->appendChild($xmlCellStyles);

        $xmlCellStylesStyle = $xml->createElement('cellStyle');
        $xmlCellStylesStyle->setAttribute('name', 'Default');
        $xmlCellStylesStyle->setAttribute('xfId', '0');
        $xmlCellStylesStyle->setAttribute('builtinId', '0');
        $xmlCellStyles->appendChild($xmlCellStylesStyle);

        /**
         * dxfs
         */
        $xmlDxfs = $xml->createElement('dxfs');
        $xmlDxfs->setAttribute('count', '0');
        $xmlStyle->appendChild($xmlDxfs);

        /**
         * Table styles
         */
        $xmlTableStyles = $xml->createElement('tableStyles');
        $xmlTableStyles->setAttribute('count', '0');
        $xmlTableStyles->setAttribute('defaultTableStyle', 'TableStyleMedium2');
        $xmlTableStyles->setAttribute('defaultPivotStyle', 'PivotStyleLight16');
        $xmlStyle->appendChild($xmlTableStyles);

        /**
         * extLst
         */
        $xmlExts = $xml->createElement('extLst');
        $xmlStyle->appendChild($xmlExts);

        $xmlExtsExt = $xml->createElement('ext');
        $xmlExtsExt->setAttribute('uri', '{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}');
        $xmlExtsExt->setAttribute('xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main');
        $xmlExts->appendChild($xmlExtsExt);

        $xmlExtsExtX14 = $xml->createElement('x14:slicerStyles');
        $xmlExtsExtX14->setAttribute('defaultSlicerStyle', 'SlicerStyleLight1');
        $xmlExtsExt->appendChild($xmlExtsExtX14);

        $xmlExtsExt2 = $xml->createElement('ext');
        $xmlExtsExt2->setAttribute('uri', '{9260A510-F301-46a8-8635-F512D64BE5F5}');
        $xmlExtsExt2->setAttribute('xmlns:x15', 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main');
        $xmlExts->appendChild($xmlExtsExt2);

        $xmlExtsExt2X15 = $xml->createElement('x15:timelineStyles');
        $xmlExtsExt2X15->setAttribute('defaultTimelineStyle', 'TimeSlicerStyleLight1');
        $xmlExtsExt2->appendChild($xmlExtsExt2X15);

        return $xml->saveXML();
    }

    /**
     * @param Style $style
     * @return int
     */
    protected function getFontId(Style $style): int
    {
        return array_search($style->getFontHash(), array_keys($this->fonts));
    }

    /**
     * @param Style $style
     * @return int
     */
    protected function getFillId(Style $style): int
    {
        return array_search($style->getFillHash(), array_keys($this->fills));
    }

    /**
     * @param Style $style
     * @return int
     */
    protected function getBorderId(Style $style): int
    {
        return array_search($style->getBorderHash(), array_keys($this->borders));
    }

}
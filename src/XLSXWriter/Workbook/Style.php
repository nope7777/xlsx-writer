<?php

namespace XLSXWriter\Workbook;

/**
 * Class Style
 */
class Style
{

    public const BORDER_STYLE_NONE = 'none';
    public const BORDER_STYLE_THIN = 'thin';
    public const BORDER_STYLE_SOLID = 'solid';
    public const BORDER_STYLE_DASHED = 'dashed';
    public const BORDER_STYLE_DOTTED = 'dotted';
    public const BORDER_STYLE_DOUBLE = 'double';

    /**
     * @var string|null
     */
    protected $horizontalAlignment;

    /**
     * @var string|null
     */
    protected $verticalAlignment;

    /**
     * @var int|null
     */
    protected $numberFormat;

    /**
     * @var string
     */
    protected $fontName = 'Calibri';

    /**
     * @var int
     */
    protected $fontSize = 11;

    /**
     * @var string|null
     */
    protected $fontColor;

    /**
     * @var bool
     */
    protected $fontBold = false;

    /**
     * @var bool
     */
    protected $fontItalic = false;

    /**
     * @var bool
     */
    protected $fontUnderline = false;

    /**
     * @var bool
     */
    protected $fontStrikethrough = false;

    /**
     * @var string
     */
    protected $fillPatternType = 'none';

    /**
     * @var string|null
     */
    protected $fillColor;

    /**
     * @var string|null
     */
    protected $borderLeftStyle;

    /**
     * @var string|null
     */
    protected $borderLeftColor;

    /**
     * @var string|null
     */
    protected $borderRightStyle;

    /**
     * @var string|null
     */
    protected $borderRightColor;

    /**
     * @var string|null
     */
    protected $borderTopStyle;

    /**
     * @var string|null
     */
    protected $borderTopColor;

    /**
     * @var string|null
     */
    protected $borderBottomStyle;

    /**
     * @var string|null
     */
    protected $borderBottomColor;

    /**
     * @var string|null
     */
    protected $borderDiagonalStyle;

    /**
     * @var string|null
     */
    protected $borderDiagonalColor;

    /**
     * @return string
     */
    public function getFontName(): string
    {
        return $this->fontName;
    }

    /**
     * @param string $fontName
     * @return Style
     */
    public function setFontName(string $fontName): self
    {
        $this->fontName = $fontName;

        return $this;
    }

    /**
     * @return int
     */
    public function getFontSize(): int
    {
        return $this->fontSize;
    }

    /**
     * @param int $fontSize
     * @return Style
     */
    public function setFontSize(int $fontSize): self
    {
        $this->fontSize = $fontSize;

        return $this;
    }

    /**
     * @return string
     */
    public function getFillPatternType(): string
    {
        return $this->fillPatternType;
    }

    /**
     * @param string $fillPatternType
     * @return Style
     */
    public function setFillPatternType(string $fillPatternType): self
    {
        $this->fillPatternType = $fillPatternType;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getHorizontalAlignment(): ?string
    {
        return $this->horizontalAlignment;
    }

    /**
     * @param string $horizontalAlignment
     * @return Style
     */
    public function setHorizontalAlignment(string $horizontalAlignment): self
    {
        $this->horizontalAlignment = $horizontalAlignment;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getVerticalAlignment(): ?string
    {
        return $this->verticalAlignment;
    }

    /**
     * @param string $verticalAlignment
     * @return Style
     */
    public function setVerticalAlignment(string $verticalAlignment): self
    {
        $this->verticalAlignment = $verticalAlignment;

        return $this;
    }

    /**
     * @return int|null
     */
    public function getNumberFormat(): ?int
    {
        return $this->numberFormat;
    }

    /**
     * @param int $numberFormat
     * @return Style
     */
    public function setNumberFormat(int $numberFormat): self
    {
        $this->numberFormat = $numberFormat;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getFontColor(): ?string
    {
        return $this->fontColor;
    }

    /**
     * @param string $fontColor
     * @return Style
     */
    public function setFontColor(string $fontColor): self
    {
        $this->fontColor = $fontColor;

        return $this;
    }

    /**
     * @return bool
     */
    public function isFontBold(): bool
    {
        return $this->fontBold;
    }

    /**
     * @param bool $fontBold
     * @return Style
     */
    public function setFontBold(bool $fontBold): self
    {
        $this->fontBold = $fontBold;

        return $this;
    }

    /**
     * @return bool
     */
    public function isFontItalic(): bool
    {
        return $this->fontItalic;
    }

    /**
     * @param bool $fontItalic
     * @return Style
     */
    public function setFontItalic(bool $fontItalic): self
    {
        $this->fontItalic = $fontItalic;

        return $this;
    }

    /**
     * @return bool
     */
    public function isFontUnderline(): bool
    {
        return $this->fontUnderline;
    }

    /**
     * @param bool $fontUnderline
     * @return Style
     */
    public function setFontUnderline(bool $fontUnderline): self
    {
        $this->fontUnderline = $fontUnderline;

        return $this;
    }

    /**
     * @return bool
     */
    public function isFontStrikethrough(): bool
    {
        return $this->fontStrikethrough;
    }

    /**
     * @param bool $fontStrikethrough
     * @return Style
     */
    public function setFontStrikethrough(bool $fontStrikethrough): self
    {
        $this->fontStrikethrough = $fontStrikethrough;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getFillColor(): ?string
    {
        return $this->fillColor;
    }

    /**
     * @param string $fillColor
     * @return Style
     */
    public function setFillColor(string $fillColor): self
    {
        $this->fillColor = $fillColor;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderLeftStyle(): ?string
    {
        return $this->borderLeftStyle;
    }

    /**
     * @param string $borderLeftStyle
     * @return Style
     */
    public function setBorderLeftStyle(string $borderLeftStyle): self
    {
        $this->borderLeftStyle = $borderLeftStyle;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderLeftColor(): ?string
    {
        return $this->borderLeftColor;
    }

    /**
     * @param string $borderLeftColor
     * @return Style
     */
    public function setBorderLeftColor(string $borderLeftColor): self
    {
        $this->borderLeftColor = $borderLeftColor;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderRightStyle(): ?string
    {
        return $this->borderRightStyle;
    }

    /**
     * @param string $borderRightStyle
     * @return Style
     */
    public function setBorderRightStyle(string $borderRightStyle): self
    {
        $this->borderRightStyle = $borderRightStyle;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderRightColor(): ?string
    {
        return $this->borderRightColor;
    }

    /**
     * @param string $borderRightColor
     * @return Style
     */
    public function setBorderRightColor(string $borderRightColor): self
    {
        $this->borderRightColor = $borderRightColor;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderTopStyle(): ?string
    {
        return $this->borderTopStyle;
    }

    /**
     * @param string $borderTopStyle
     * @return Style
     */
    public function setBorderTopStyle(string $borderTopStyle): self
    {
        $this->borderTopStyle = $borderTopStyle;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderTopColor(): ?string
    {
        return $this->borderTopColor;
    }

    /**
     * @param string $borderTopColor
     * @return Style
     */
    public function setBorderTopColor(string $borderTopColor): self
    {
        $this->borderTopColor = $borderTopColor;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderBottomStyle(): ?string
    {
        return $this->borderBottomStyle;
    }

    /**
     * @param string $borderBottomStyle
     * @return Style
     */
    public function setBorderBottomStyle(string $borderBottomStyle): self
    {
        $this->borderBottomStyle = $borderBottomStyle;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderBottomColor(): ?string
    {
        return $this->borderBottomColor;
    }

    /**
     * @param string $borderBottomColor
     * @return Style
     */
    public function setBorderBottomColor(string $borderBottomColor): self
    {
        $this->borderBottomColor = $borderBottomColor;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderDiagonalStyle(): ?string
    {
        return $this->borderDiagonalStyle;
    }

    /**
     * @param string $borderDiagonalStyle
     * @return Style
     */
    public function setBorderDiagonalStyle(string $borderDiagonalStyle): self
    {
        $this->borderDiagonalStyle = $borderDiagonalStyle;

        return $this;
    }

    /**
     * @return string|null
     */
    public function getBorderDiagonalColor(): ?string
    {
        return $this->borderDiagonalColor;
    }

    /**
     * @param string $borderDiagonalColor
     * @return Style
     */
    public function setBorderDiagonalColor(string $borderDiagonalColor): self
    {
        $this->borderDiagonalColor = $borderDiagonalColor;

        return $this;
    }

    /**
     * @param string $left
     * @param string $top
     * @param string $right
     * @param string $bottom
     */
    public function setBordersStyles(string $left, string $top, string $right, string $bottom): void
    {
        $this->borderLeftStyle = $left;
        $this->borderTopStyle = $top;
        $this->borderRightStyle = $right;
        $this->borderBottomStyle = $bottom;
    }

    /**
     * @param string $left
     * @param string $top
     * @param string $right
     * @param string $bottom
     */
    public function setBordersColors(string $left, string $top, string $right, string $bottom): void
    {
        $this->borderLeftColor = $left;
        $this->borderTopColor = $top;
        $this->borderRightColor = $right;
        $this->borderBottomColor = $bottom;
    }

    /**
     * @param string $style
     * @param string $color
     */
    public function setBorders(string $style, string $color): void
    {
        $this->borderLeftStyle = $style;
        $this->borderTopStyle = $style;
        $this->borderRightStyle = $style;
        $this->borderBottomStyle = $style;

        $this->borderLeftColor = $color;
        $this->borderTopColor = $color;
        $this->borderRightColor = $color;
        $this->borderBottomColor = $color;
    }

    /**
     * @param array $values
     * @return string
     */
    protected function getValuesHash(array $values): string
    {
        $values = implode(':', $values);

        return md5($values);
    }

    /**
     * @return string
     */
    public function getFillHash(): string
    {
        return $this->getValuesHash([
            $this->fillPatternType,
            $this->fillColor,
        ]);
    }

    /**
     * @return string
     */
    public function getFontHash(): string
    {
        return $this->getValuesHash([
            $this->fontSize,
            $this->fontName,
            $this->fontColor,
            $this->fontBold,
            $this->fontItalic,
            $this->fontUnderline,
            $this->fontStrikethrough,
        ]);
    }

    /**
     * @return string
     */
    public function getBorderHash(): string
    {
        return $this->getValuesHash([
            $this->borderLeftStyle,
            $this->borderLeftColor,
            $this->borderRightStyle,
            $this->borderRightColor,
            $this->borderTopStyle,
            $this->borderTopColor,
            $this->borderBottomStyle,
            $this->borderBottomColor,
            $this->borderDiagonalStyle,
            $this->borderDiagonalColor,
        ]);
    }

    /**
     * @return string
     */
    public function getHash(): string
    {
        return $this->getValuesHash([
            $this->getFillHash(),
            $this->getFontHash(),
            $this->getBorderHash(),
        ]);
    }

}
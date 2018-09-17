<?php

namespace XLSXWriter\Helper;

use DateTime;

/**
 * Class ExcelHelper
 */
class ExcelHelper
{

    /**
     * @var array
     */
    protected static $indexCache = [];

    /**
     * @var string
     */
    protected static $escapableControlCharactersPattern = '[\x00-\x08\x0B-\x0C\x0E-\x1F]';

    /**
     * Converting column index (int) to letter (string).
     * Starts with 0 (0 = A).
     *
     * @param int $index
     * @return string
     */
    public static function indexToLetter(int $index): string
    {
        if (!isset(self::$indexCache[$index])) {
            // Determine column string
            if ($index < 26) {
                self::$indexCache[$index] = chr(65 + $index);
            } elseif ($index < 702) {
                self::$indexCache[$index] = chr(64 + ($index / 26)) .
                    chr(65 + $index % 26);
            } else {
                self::$indexCache[$index] = chr(64 + (($index - 26) / 676)) .
                    chr(65 + ((($index - 26) % 676) / 26)) .
                    chr(65 + $index % 26);
            }
        }
        return self::$indexCache[$index];
    }

    /**
     * @param string $string
     * @return string
     */
    public static function escapeString(string $string): string
    {
        $escapedString = self::escapeControlCharacters($string);

        $escapedString = htmlspecialchars($escapedString, ENT_NOQUOTES);

        return $escapedString;
    }

    /**
     * @param string $string
     * @return string
     */
    public static function escapeControlCharacters(string $string): string
    {
        $escapedString = self::escapeEscapeCharacter($string);

        // if no control characters
        if (!preg_match("/{" . self::$escapableControlCharactersPattern . "}/", $escapedString)) {
            return $escapedString;
        }

        return $escapedString;
    }

    /**
     * @param string $string
     * @return string
     */
    public static function escapeEscapeCharacter(string $string): string
    {
        return preg_replace('/_(x[\dA-F]{4})_/', '_x005F_$1_', $string);
    }

    /**
     * @param int|DateTime $dateValue
     * @return bool|float
     */
    public static function PHPToExcel($dateValue = 0)
    {
        $saveTimeZone = date_default_timezone_get();
        date_default_timezone_set('UTC');
        $retValue = FALSE;
        if ($dateValue instanceof DateTime) {
            $retValue = self::FormattedPHPToExcel( $dateValue->format('Y'), $dateValue->format('m'), $dateValue->format('d'),
                $dateValue->format('H'), $dateValue->format('i'), $dateValue->format('s')
            );
        }
        elseif (is_numeric($dateValue)) {
            $retValue = self::FormattedPHPToExcel( date('Y',$dateValue), date('m',$dateValue), date('d',$dateValue),
                date('H',$dateValue), date('i',$dateValue), date('s',$dateValue)
            );
        }
        date_default_timezone_set($saveTimeZone);

        return $retValue;
    }

    /**
     * @param int $year
     * @param int $month
     * @param int $day
     * @param int $hours
     * @param int $minutes
     * @param int $seconds
     * @return float
     */
    public static function FormattedPHPToExcel(int $year, int $month, int $day, int $hours = 0, int $minutes = 0, int $seconds = 0): float
    {
        //
        //	Fudge factor for the erroneous fact that the year 1900 is treated as a Leap Year in MS Excel
        //	This affects every date following 28th February 1900
        //
        $excel1900isLeapYear = TRUE;
        if (($year == 1900) && ($month <= 2)) { $excel1900isLeapYear = FALSE; }
        $my_excelBaseDate = 2415020;

        //	Julian base date Adjustment
        if ($month > 2) {
            $month -= 3;
        } else {
            $month += 9;
            --$year;
        }

        //	Calculate the Julian Date, then subtract the Excel base date (JD 2415020 = 31-Dec-1899 Giving Excel Date of 0)
        $century = substr($year,0,2);
        $decade = substr($year,2,2);
        $excelDate = floor((146097 * $century) / 4) + floor((1461 * $decade) / 4) + floor((153 * $month + 2) / 5) + $day + 1721119 - $my_excelBaseDate + $excel1900isLeapYear;

        $excelTime = (($hours * 3600) + ($minutes * 60) + $seconds) / 86400;

        return (float) $excelDate + $excelTime;
    }

}
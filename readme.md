## About XLSXWriter

XLSXWriter is lightweight library for generating large Excel exports fast and without memory overflow.

## Documentation

Usage example:

```php
use XLSXWriter\Workbook;

$workbook = new Workbook();

$workbook->getDefaultStyle()
    ->setFontName('Arial')
    ->setFontSize(9);

$sheet = $workbook->addSheet('Sheet title');
$sheet->addRow(
    ['Person', 'Company', 'Amount'],
    $workbook->getNewStyle()->setFontBold(true)
);
$sheet->skipRows(2);

$sheet->addRows([
    ['User1', 'Company1', 100],
    ['User2', 'Company2', 200],
    ['User3', 'Company3', 300],
]);

$workbook->output('some-export.xlsx');
```



# Simple Excel exporter

Simple library which allows to quickly create MS Excel exports.

Key features:
 - Simple
 - Multi sheet support
 - Values can be accessed by the [PropertyAccess Component](https://symfony.com/doc/current/components/property_access.html) or [PHP callback](https://www.php.net/manual/en/language.types.callable.php).
 - XLS, XLSX and CSV support (by passing the second argument while creating a new \SecIT\SimpleExcelExport\Excel obejct)
 - Symfony 4+ compatible (getResponse() returns valid [Symfony Response](https://symfony.com/doc/current/components/http_foundation.html#response))
 
## Installation

From the command line run

```
$ composer require secit-pl/simple-excel-export
```

## Usage

### Basic example

#### Send Excel file as response to user

```php
<?php

use SecIT\SimpleExcelExport\Excel;

// example data
$data = [
    'Simple array example' => [
        ['col1' => 123, 'col2' => 321],
        ['col1' => 234, 'col2' => 345],
    ],
];

$excel = new Excel('test', Excel::OUTPUT_XLSX);
$excel->setColumnsAutoSizingEnabled(true);

$excel->addSheet('Simple array example')
    ->setColumn('Column 1', '[col1]') // use Symfony property access component notation or callback
    ->setColumn('Column 2', '[col2]');

// get response (Symfony compatible) 
$response = $excel->getResponse($data)

// and sent it to the browser
$response->send();
```

#### Create Excel file 

```php
<?php

use SecIT\SimpleExcelExport\Excel;

// example data
$data = [
    'Simple array example' => [
        ['col1' => 123, 'col2' => 321],
        ['col1' => 234, 'col2' => 345],
    ],
];

$excel = new Excel('test', Excel::OUTPUT_XLSX);
$excel->setColumnsAutoSizingEnabled(true);

$excel->addSheet('Simple array example')
    ->setColumn('Column 1', '[col1]') // use Symfony property access component notation or callback
    ->setColumn('Column 2', '[col2]');

// get file 
$splFileObject = $excel->getFile('/path/to/the/file.xlsx', $data);
```

### Advanced example

```php
<?php

use SecIT\SimpleExcelExport\Excel;

// Excel data
// data class used in this example
class ExampleUser {
    public $name;
    public $surname;
    public $parent;

    public function __construct($name, $surname, ExampleUser $parent = null) {
        $this->name = $name;
        $this->surname = $surname;
        $this->parent = $parent;
    }
}

// the data
$data = [
    'Simple array example' => [
        ['col1' => 123, 'col2' => 321],
        ['col1' => 234, 'col2' => 345],
    ],
    'Filters example' => [
        ['col3' => 'So sad', 'col4' => new \DateTime()],
        ['col3' => 'So happy', 'col4' => new \DateTime('1234-12-11 11:11:22')],
    ],
    'Objects example' => [
        new ExampleUser('John', 'Blue', new ExampleUser('Jan', 'Blue')),
        new ExampleUser('Jack', 'Red', new ExampleUser('Tom', 'Red')),
    ],
    'Callback example' => [
        ['col1' => 1, 'col2' => 2, 'col3' => null],
        ['col1' => 3, 'col2' => 4, 'col3' => null],
    ],
];

// Create the new Excel object
$excel = new Excel('test', Excel::OUTPUT_XLSX);
$excel->setColumnsAutoSizingEnabled(true);

// Simple array example
$excel->addSheet('Simple array example')
    ->setColumn('Column 1', '[col1]')
    ->setColumn('Column 2', '[col2]');

// Filters example
$excel->addSheet('Filters example')
    ->setColumn('Column 3', '[col3]', [
        new Excel\Filter\PregReplaceFilter('/sad/', 'happy'),
    ])
    ->setColumn('Column 4', '[col4]', [
        new Excel\Filter\DateTimeFilter('d.m.Y'),
    ]);

// Objects example
$excel->addSheet('Objects example')
    ->setColumn('Name', 'name')
    ->setColumn('Surname', 'surname')
    ->setColumn('Parent name', 'parent.name')
;

// Callback example
$excel->addSheet('Callback example')
    ->setColumn('Column 1', '[col1]')
    ->setColumn('Column 2', '[col2]')
    ->setColumn('Column 1 + Column 2', static function ($row) {
        return $row['col1'] + $row['col2'];
    });

// Get response and sent it to the browser
$excel->getResponse($data)
    ->send();
```
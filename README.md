# Edit an Excel file that contains macros for PHP
This PHP library allows you to modify an Excel file that contains macros without disabling them.
For most PHP projects, the library (https://github.com/PHPOffice/PhpSpreadsheet) allows you to manipulate Excel files.
However, when you modify an Excel file with macro (xlsm), the macros no longer work.
This PHP library will allow you to modify XLSM files without disabling macros.
The library only applies to Excel files with macros

Requirements
-------------------
* [PHP 7.4.0 or higher with CURL Support](http://www.php.net/)

Installation
-------------------
You can use **Composer** or simply **Download the Release**

#### Composer
The preferred method is via [composer](https://getcomposer.org). Follow the
[installation instructions](https://getcomposer.org/doc/00-intro.md) if you do not already have
composer installed.

Once composer is installed, execute the following command in your project root to install this library:

```sh
  composer require ribojhin/excelmacro
```

Finally, be sure to include the autoloader:

```php
<?php
  require_once '/path/to/your-project/vendor/autoload.php';
```

Quickstart
-------------------
The example below will load data into an excel file that contains macros:
```php
<?php
  require_once 'autoload.php';                                      // Comment this line if you use Composer to install the package
  use \Ribojhin\Excelmacro;

  $excelMacro = new Excelmacro("_FILE_PATH_SRC_XLSM_");              // Set excel file path you want to edit
  $excelMacro->setSheet(0);                                         // Set sheet by index
  $excelMacro->setCellValue($excelMacro->getSheet(), "_KEY_", '_VALUE_'); // Set cell
  $excelMacro->save("_FILE_PATH_DEST_XLSM_");                        // Set destination file path

```
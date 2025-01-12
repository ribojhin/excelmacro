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
  composer require convertio/convertio-php
```

Finally, be sure to include the autoloader:

```php
<?php
  require_once '/path/to/your-project/vendor/autoload.php';
```

Quickstart
-------------------
Following example will render remote web page into PNG image:
```php
<?php
  require_once 'autoload.php';                      // Comment this line if you use Composer to install the package
  use \Convertio\Convertio;

  $API = new Convertio("_YOUR_API_KEY_");           // You can obtain API Key here: https://convertio.co/api/

  $API->startFromURL('http://google.com/', 'png')   // Convert (Render) HTML Page to PNG
  ->wait()                                          // Wait for conversion finish
  ->download('./google.png')                        // Download Result To Local File
  ->delete();                                       // Delete Files from Convertio hosts
```
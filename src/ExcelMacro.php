<?php
declare(strict_types=1);

namespace Ribojhin\ExcelmacroBundle;

use Ribojhin\ExcelmacroBundle\Sheet;
use Ribojhin\ExcelmacroBundle\Archive;

class ExcelMacro
{
    private string $root;

    private Archive $archive;

    private Sheet $sheet;

    public function __construct(string $filePath)
    {
        $ext = pathinfo($filePath, PATHINFO_EXTENSION);
        if ($ext !== DataType::EXT_XLSM) throw new \Exception(sprintf('Invalid extension file (%s)', $ext));
        
        $this->root = sys_get_temp_dir();
        $this->archive = new Archive();
        // convert xlsm to zip
        $this->archive->convertXlsmToZip($this->root, $filePath);
    }

    /**
     * @return Sheet
     */
    public function getSheet()
    {
        return $this->sheet;
    }

    /**
     * @var int $index
     * 
     * @return Sheet
     */
    public function setSheet(int $index)
    {
        $sheet = new Sheet();
        $sheet->setIndex($index);
        $sheet->setSharedFilePath($this->root . DIRECTORY_SEPARATOR . $this->archive->getZipname() . DIRECTORY_SEPARATOR . 'xl' . DIRECTORY_SEPARATOR . 'sharedStrings.xml');
        $sheet->setFilePath($this->root . DIRECTORY_SEPARATOR . $this->archive->getZipname() . DIRECTORY_SEPARATOR . 'xl' . DIRECTORY_SEPARATOR . 'worksheets' . DIRECTORY_SEPARATOR . 'sheet' . ($index + 1) . '.' . DataType::EXT_XML);
        $this->sheet = $sheet;

        return $this;
    }

    /**
     * @var Sheet $sheet
     * @var string $key
     * @var string $value
     * @var string $typeData
     */
    public function setCellValue(Sheet $sheet, string $key, mixed $value, string $typeData = DataType::TYPE_STRING): void
    {
        $newData = [];
        if (self::isString($value) || $typeData == DataType::TYPE_STRING) {
            $value = strval($value);
            $document = new \DOMDocument();
            $document->loadXml(file_get_contents($sheet->getSharedFilePath()));
            $sst = $document->getElementsByTagName('sst')->item(0);
            
            $newElement = $document->createElement("si");
            $sst->appendChild($newElement);
            $elements = $document->getElementsByTagName('si')->length;
            $si = $document->getElementsByTagName('si')->item($elements - 1);
            $newElement = $document->createElement('t');
            $newElement->textContent = $value;
            $si->appendChild($newElement);
            $newData[$key] = strval($document->getElementsByTagName('si')->length - 1);
            
            $sst->setAttribute('count', strval($sst->getAttribute('count') + 1));
            $sst->setAttribute('uniqueCount', strval($document->getElementsByTagName('si')->length));
            $document->save($sheet->getSharedFilePath());
        }
        //
        $document = new \DOMDocument();
        $document->loadXml(file_get_contents($sheet->getFilePath()));
        $cs = $document->getElementsByTagName('c');
        foreach ($cs as $c) {
            if ($c->getAttribute('r') == $key) {
                $child = $document->createElement('v');
                if (array_key_exists($key, $newData)) {
                    $child->textContent = strval($newData[$key]);
                    $c->appendChild($child);
                    $c->setAttribute('t', 's');
                } else {
                    $child->textContent = strval($value);
                    $c->appendChild($child);
                }
                break;
            }
        }
        $document->save($sheet->getFilePath());
    }

    /**
     * @var Sheet $sheet
     * @var string $key
     * 
     * @return mixed
     */
    public function getCellValue(Sheet $sheet, string $key): mixed
    {
        $value = '';
        $hasSharedData = False;
        //
        $document = new \DOMDocument();
        $document->loadXml(file_get_contents($sheet->getFilePath()));
        $cs = $document->getElementsByTagName('c');
        foreach ($cs as $c) {
            if ($c->getAttribute('r') == $key) {
                if($c->hasAttribute('t')) {
                    $hasSharedData = True;
                }
                $values = $c->childNodes;
                if ($values->length > 0) {
                    $value = $values->item(0);
                }
                break;
            }
        }
        if ($hasSharedData && is_numeric($value)) {
            $document = new \DOMDocument();
            $document->loadXml(file_get_contents($sheet->getSharedFilePath()));
            try {
                $si = $document->getElementsByTagName('si')->item(intval($value));
                $value = $si->nodeValue;
            } catch (\Exception $ex) {
                $value = '';
            }
        }
        return $value;
    }

    /**
     * @var string $destinationFilePath
     */
    public function save(string $destinationFilePath)
    {
        $ext = pathinfo($destinationFilePath, PATHINFO_EXTENSION);
        if ($ext !== DataType::EXT_XLSM) throw new \Exception(sprintf('Invalid extension file (%s)', $ext));
        // convert zip to xlsm
        $this->archive->convertZipToXlsm($this->root, $destinationFilePath);
    }

    public static function isDecimalOrInt($value)
    {
        return (preg_match('/^\d+\.\d+$/',$value) || preg_match('/^\d+$/',$value));
    }

    public static function isString($value)
    {
        if (is_float($value)) {
            return false;
        }
        if (self::isDecimalOrInt($value)) {
            return false;
        }
        return true;
    }
}
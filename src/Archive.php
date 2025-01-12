<?php
declare(strict_types=1);

namespace Ribojhin\Excelmacro;

class Archive
{
    private string $zipname;

    public function __construct() {

    }

    /**
     * @return string
     */
    public function getZipname()
    {
        return $this->zipname;
    }

    /**
     * @var string $zipname
     * 
     * @return Archive
     */
    public function setZipname(string $zipname)
    {
        $this->zipname = $zipname;

        return $this;
    }

    /**
     * @var string $root
     * @var string $xlsmFilePath
     */
    public function convertXlsmToZip(string $root, string $xlsmFilePath): void
    {
        $destZipName = date('YmdHms').uniqid();
        $destXlsmFileName = $destZipName . '.' . DataType::EXT_XLSM;
        copy($xlsmFilePath, $root . DIRECTORY_SEPARATOR . $destXlsmFileName);
        // rename file
        rename($root . DIRECTORY_SEPARATOR . $destXlsmFileName, $root . DIRECTORY_SEPARATOR . $destZipName . '.' . DataType::EXT_ZIP);
        $zip = new \ZipArchive;
        $res = $zip->open($root . DIRECTORY_SEPARATOR . $destZipName . '.' . DataType::EXT_ZIP);
        if ($res === TRUE) {
            $zip->extractTo($root . DIRECTORY_SEPARATOR . $destZipName);
            $zip->close();
            // Delete zip file
            unlink($root . DIRECTORY_SEPARATOR . $destZipName . '.' . DataType::EXT_ZIP);
        }
        $this->setZipname($destZipName);
    }

    /**
     * @var string $root
     * @var string $destinationFilePath
     * @var bool $isDelete
     */
    public function convertZipToXlsm(string $root, string $destinationFilePath, bool $isDelete = True): void
    {
        $destination = $root . DIRECTORY_SEPARATOR . $this->zipname . '.' . DataType::EXT_XLSM;
        // zip folders
        $this->Zip($root . DIRECTORY_SEPARATOR . $this->zipname, $root . DIRECTORY_SEPARATOR . $this->zipname .'.' . DataType::EXT_ZIP);
        if (!empty($destinationFilePath) && pathinfo($destinationFilePath, PATHINFO_EXTENSION) == DataType::EXT_XLSM) {
            rename($root . DIRECTORY_SEPARATOR . $this->zipname . '.' . DataType::EXT_ZIP, $destinationFilePath);
        }
        if ($isDelete) {
            // Delete zip folder
            $this->deleteFolder($root . DIRECTORY_SEPARATOR . $this->zipname);
        }
    }

    /**
     * @var string $source
     * @var string $destination
     * 
     * @return bool
     */
    public function Zip(string $source, string $destination): bool
    {
        if (!extension_loaded(DataType::EXT_ZIP) || !file_exists($source)) {
            return false;
        }
    
        $zip = new \ZipArchive();
        if (!$zip->open($destination, \ZIPARCHIVE::CREATE)) {
            return false;
        }
        if ($source instanceof \SplFileInfo) {
            $source = $source->getPathname();
        }
        $source = str_replace('\\', '/', realpath($source));
        if (is_dir($source) === true) {
            $files = new \RecursiveIteratorIterator(new \RecursiveDirectoryIterator($source), \RecursiveIteratorIterator::SELF_FIRST);
    
            foreach ($files as $file)
            {
                if ($file instanceof \SplFileInfo) {
                    $file = $file->getPathname();
                }
                $file = str_replace('\\', '/', $file);
                if( in_array(substr($file, strrpos($file, '/')+1), array('.', '..')) )
                    continue;
                // $file = realpath($file);
                if (is_dir($file) === true) {
                    $zip->addEmptyDir(str_replace($source . '/', '', $file . '/'));
                } else if (is_file($file) === true) {
                    $zip->addFromString(str_replace($source . '/', '', $file), file_get_contents($file));
                }
            }
        } else if (is_file($source) === true) {
            $zip->addFromString(basename($source), file_get_contents($source));
        }
        return $zip->close();
    }

    /**
     * @var string $source
     */
    public function deleteFolder(string $source)
    {
        $dir = opendir($source);
        while(false !== ( $file = readdir($dir)) ) {
            if (( $file != '.' ) && ( $file != '..' )) {
                $full = $source . DIRECTORY_SEPARATOR . $file;
                if ( is_dir($full) ) {
                    $this->deleteFolder($full);
                } else {
                    unlink($full);
                }
            }
        }
        closedir($dir);
        rmdir($source);
    }
}
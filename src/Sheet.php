<?php
namespace Ribojhin\Excelmacro;

class Sheet
{
    private int $index;

    private string $filePath;

    private string $sharedFilePath;

    public function __construct() {

    }

    /**
     * @return int
     */
    public function getIndex(): int
    {
        return $this->index;
    }

    /**
     * @var int $index
     * 
     * @return Sheet
     */
    public function setIndex(int $index):self
    {
        $this->index = $index;

        return $this;
    }

    /**
     * @return string
     */
    public function getFilePath(): string
    {
        return $this->filePath;
    }

    /**
     * @var string $filePath
     * 
     * @return Sheet
     */
    public function setFilePath(string $filePath): self
    {
        if (!file_exists($filePath)) throw new \Exception(sprintf('file %s doesn\'t exist', basename($filePath)));
        $this->filePath = $filePath;

        return $this;
    }

    /**
     * @return string
     */
    public function getSharedFilePath(): string
    {
        return $this->sharedFilePath;
    }

    /**
     * @var string $sharedFilePath
     * 
     * @return Sheet
     */
    public function setSharedFilePath(string $sharedFilePath): self
    {
        if (!file_exists($sharedFilePath)) throw new \Exception(sprintf('file %s doesn\'t exist', basename($sharedFilePath)));
        $this->sharedFilePath = $sharedFilePath;

        return $this;
    }
}
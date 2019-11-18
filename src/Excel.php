<?php

declare(strict_types=1);

namespace SecIT\SimpleExcelExport;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\BaseWriter;
use PhpOffice\PhpSpreadsheet\Writer\Csv;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SecIT\SimpleExcelExport\Excel\Sheet;
use Symfony\Component\HttpFoundation\Response;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;
use Symfony\Component\HttpFoundation\StreamedResponse;

class Excel
{
    public const OUTPUT_XLSX = 'xlsx';
    public const OUTPUT_XLS = 'xls';
    public const OUTPUT_CSV = 'csv';

    private $fileName;
    private $outputFormat;

    /**
     * @var Sheet[]
     */
    private $sheets;

    private const OUTPUT_WRITERS = [
        self::OUTPUT_XLSX => Xlsx::class,
        self::OUTPUT_XLS => Xls::class,
        self::OUTPUT_CSV => Csv::class,
    ];

    private const RESPONSE_CONTENT_TYPES = [
        self::OUTPUT_XLSX => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        self::OUTPUT_XLS => 'application/vnd.ms-excel',
        self::OUTPUT_CSV => 'application/csv',
    ];

    public function __construct(string $fileName = 'Excel', string $outputFormat = self::OUTPUT_XLSX)
    {
        $this->setFileName($fileName);
        $this->setOutputFormat($outputFormat);
    }

    public function getFileName(): string
    {
        return $this->fileName;
    }

    public function setFileName(string $fileName): Excel
    {
        $this->fileName = $fileName;

        return $this;
    }

    public function getOutputFormat(): string
    {
        return $this->outputFormat;
    }

    public function setOutputFormat(string $outputFormat): Excel
    {
        if (!isset(self::OUTPUT_WRITERS[$outputFormat])) {
            throw new \InvalidArgumentException(sprintf(
                'Invalid output format `%s`. Expected: %s.',
                $outputFormat,
                implode(array_keys(self::OUTPUT_WRITERS))
            ));
        }

        $this->outputFormat = $outputFormat;

        return $this;
    }

    public function addSheet(string $name): Sheet
    {
        $sheet = new Sheet($name);

        if ($this->hasSheet($name)) {
            throw new \InvalidArgumentException(sprintf('Sheet `%s` already exists.', $name));
        }

        $this->sheets[$name] = $sheet;

        return $sheet;
    }

    public function getSheet(string $name): Sheet
    {
        if (!$this->hasSheet($name)) {
            throw new \InvalidArgumentException(sprintf('Sheet `%s` not found.', $name));
        }

        return $this->sheets[$name];
    }

    public function hasSheet(string $name): bool
    {
        return isset($this->sheets[$name]);
    }

    public function removeSheet(string $name): Excel
    {
        if (!$this->hasSheet($name)) {
            throw new \InvalidArgumentException(sprintf('Sheet `%s` not found.', $name));
        }

        unset ($this->sheets[$name]);

        return $this;
    }

    public function getSpreadsheet(array $data): Spreadsheet
    {
        $spreadsheet = new Spreadsheet();
        $spreadsheet->removeSheetByIndex(0);

        $index = 0;
        foreach ($this->sheets as $sheet) {
            $spreadsheet->addSheet($sheet->getWorksheet($data[$sheet->getName()] ?? []), $index++);
        }

        return $spreadsheet;
    }

    public function getResponse(array $data): Response
    {
        $writer = $this->getOutputWriter($this->getSpreadsheet($data));

        $response = new StreamedResponse(
            static function () use ($writer) {
                $writer->save('php://output');
            }
        );

        $disposition = $response->headers->makeDisposition(
            ResponseHeaderBag::DISPOSITION_ATTACHMENT,
            $this->getFileName().'.'.$this->getOutputFormat()
        );

        $response->headers->set('Content-Type', self::RESPONSE_CONTENT_TYPES[$this->getOutputFormat()]);
        $response->headers->set('Content-Disposition', $disposition);
        $response->headers->set('Cache-Control','max-age=0');

        return $response;
    }

    private function getOutputWriter(Spreadsheet $spreadsheet): BaseWriter
    {
        $class = self::OUTPUT_WRITERS[$this->getOutputFormat()];

        return new $class($spreadsheet);
    }
}

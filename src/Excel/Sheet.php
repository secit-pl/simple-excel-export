<?php

declare(strict_types=1);

namespace SecIT\SimpleExcelExport\Excel;

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use SecIT\SimpleExcelExport\Excel\Filter\ColumnFilterInterface;
use Symfony\Component\PropertyAccess\PropertyAccess;

class Sheet
{
    private $name;

    /**
     * @var array|callable[]
     */
    private $columns = [];

    private $propertyAccessor;

    private $includeHeader = true;

    public function __construct(string $name)
    {
        $this->name = $name;
    }

    public function getName(): string
    {
        return $this->name;
    }

    /**
     * @param string                                                                        $name
     * @param string|callable                                                               $getter
     * @param ColumnFilterInterface|callable|array|ColumnFilterInterface[]|callable[]||null $filters
     *
     * @return Sheet
     */
    public function setColumn(string $name, $getter, $filters = null, ?string $dataType = null): self
    {
        if (!is_string($getter) && !is_callable($getter)) {
            throw new \InvalidArgumentException(sprintf(
                'Getter should be a string or callable %s given.',
                gettype($getter)
            ));
        }

        if (null !== $filters) {
            if (!is_array($filters) || is_callable($filters)) {
                $filters = [$filters];
            }

            foreach ($filters as $filter) {
                if (!$filter instanceof ColumnFilterInterface && !is_callable($filter)) {
                    throw new \InvalidArgumentException(sprintf(
                        'Filter should be an instance of %s or callable %s given.',
                        ColumnFilterInterface::class,
                        gettype($getter)
                    ));
                }
            }
        } else {
            $filters = [];
        }

        if (null !== $dataType && !in_array($dataType, $this->getAllowedDataTypes(), true)) {
            throw new \InvalidArgumentException(sprintf(
                'Invalid data type. Expected null or one of %s::TYPE_* given.',
                DataType::class
            ));
        }

        $this->columns[$name] = new class($getter, $filters, $dataType) {
            private $getter;
            private $filters;
            private $dataType;
            private $propertyAccessor;

            public function __construct($getter, array $filters, $dataType)
            {
                $this->getter = $getter;
                $this->filters = $filters;
                $this->dataType = $dataType;
                $this->propertyAccessor = PropertyAccess::createPropertyAccessor();
            }

            public function get($data)
            {
                if (is_callable($this->getter)) {
                    $value = ($this->getter)($data);
                } else {
                    $value = $this->propertyAccessor->getValue($data, $this->getter);
                }

                foreach ($this->filters as $filter) {
                    if (is_callable($filter)) {
                        $value = $filter($value);
                    } else {
                        $value = $filter->filter($value);
                    }
                }

                return $value;
            }

            public function getDataType()
            {
                return $this->dataType;
            }
        };

        return $this;
    }

    public function setIncludeHeader($includeHeader = true): self
    {
        $this->includeHeader = $includeHeader;

        return $this;
    }

    /**
     * @internal
     */
    public function getWorksheet(array $data): Worksheet
    {
        $worksheet = new Worksheet(null, $this->getName());

        [$columnIndex, $rowIndex] = Coordinate::coordinateFromString('A1');

        if ($this->includeHeader) {
            foreach (array_keys($this->columns) as $name) {
                if (null !== $name) {
                    $worksheet->getCell($columnIndex.$rowIndex)
                        ->setValue($name);
                }

                ++$columnIndex;
            }

            ++$rowIndex;
            $columnIndex = 'A';
        }

        foreach ($data as $rowData) {
            foreach ($this->columns as $column) {
                $cellData = $column->get($rowData);
                $cellDataType = $column->getDataType();

                if (null === $cellDataType) {
                    $worksheet->setCellValue($columnIndex.$rowIndex, $cellData);
                } else {
                    $worksheet->setCellValueExplicit($columnIndex.$rowIndex, $cellData, $cellDataType);
                }

                ++$columnIndex;
            }

            ++$rowIndex;
            $columnIndex = 'A';
        }

        return $worksheet;
    }

    private function getAllowedDataTypes(): array
    {
        $types = [null];
        $prefix = 'TYPE_';

        $reflectionClass = new \ReflectionClass(DataType::class);
        foreach ($reflectionClass->getConstants() as $name => $value) {
            if (strpos($name, $prefix) === 0) {
                $types[] = $value;
            }
        }

        return $types;
    }
}

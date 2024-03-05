<?php

declare(strict_types=1);

namespace SecIT\SimpleExcelExport\Excel;

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
    public function setColumn(string $name, $getter, $filters = null): self
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

        $this->columns[$name] = new class($getter, $filters) {
            private $getter;
            private $filters;
            private $propertyAccessor;

            public function __construct($getter, array $filters)
            {
                $this->getter = $getter;
                $this->filters = $filters;
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
        };

        return $this;
    }

    /**
     * @internal
     */
    public function getWorksheet(array $data): Worksheet
    {
        $worksheet = new Worksheet(null, $this->getName());

        $headers = [array_keys($this->columns)];
        $dataRows = array_map(function ($row) {
            $rowData = [];
            foreach ($this->columns as $name => $column) {
                $rowData[] = $column->get($row);
            }

            return $rowData;
        }, $data);

        $worksheet->fromArray(array_merge($headers, $dataRows), null, 'A1', true);

        return $worksheet;
    }
}

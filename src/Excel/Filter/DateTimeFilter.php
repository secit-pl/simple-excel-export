<?php

declare(strict_types=1);

namespace SecIT\SimpleExcelExport\Excel\Filter;

class DateTimeFilter implements ColumnFilterInterface
{
    private $format;

    public function __construct(string $format = 'Y-m-d H:i:s')
    {
        $this->format = $format;
    }

    public function filter($value)
    {
        if ($value instanceof \DateTimeInterface) {
            return $value->format($this->format);
        }

        return $value;
    }
}

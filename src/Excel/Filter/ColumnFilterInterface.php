<?php

declare(strict_types=1);

namespace SecIT\SimpleExcelExport\Excel\Filter;

interface ColumnFilterInterface
{
    /**
     * @param mixed $value
     *
     * @return mixed
     */
    public function filter($value);
}

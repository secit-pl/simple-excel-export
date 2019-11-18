<?php

declare(strict_types=1);

namespace SecIT\SimpleExcelExport\Excel\Filter;

class PregReplaceFilter implements ColumnFilterInterface
{
    private $pattern;
    private $replacement;
    private $limit;

    public function __construct(string $pattern, string $replacement, int $limit = -1)
    {
        $this->pattern = $pattern;
        $this->replacement = $replacement;
        $this->limit = $limit;
    }

    public function filter($value)
    {
        return preg_replace($this->pattern, $this->replacement, $value, $this->limit);
    }
}

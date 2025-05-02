<?php

namespace App\ExcelPlatMerah;

use PhpOffice\PhpSpreadsheet\Style\Alignment as StyleAlignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

class ExcelPlatMerahStyle
{
    protected $countData, $startRowHeader, $endRowHeader, $startRowData, $startColumnUsed, $endColumnUsed, $lastRowData;

    protected $mainHeaderMergeToCenter = [];
    protected $mergeCellToBottom = [];
    protected $mergeCellToRight = [];
    protected $headerToCenterAlign = [];
    protected $columnDataToLeftAlign = [];
    protected $columnDataToCenterAlign = [];
    protected $columnDataToRightAlign = [];
    protected $rowDataBold = [];
    protected $cellDataAlign = [];
    protected $cellBackgroundColor = [];
    protected $cellBorder = [];

    public function __construct($countData, $startRowHeader, $endRowHeader, $startRowData, $startColumnUsed, $endColumnUsed)
    {
        $this->startRowHeader = $startRowHeader;
        $this->endRowHeader = $endRowHeader;
        $this->startRowData = $startRowData;
        $this->startColumnUsed = $startColumnUsed;
        $this->endColumnUsed = $endColumnUsed;
        $this->countData = $countData;
        $this->lastRowData = $this->countData + ($this->startRowData - 1);
    }

    public function applyStyle($sheet)
    {
        $styles = $this->getStyle();

        if (!empty($styles['mergeCells'])) {
            foreach ($styles['mergeCells'] as $merge) {
                if (isset($merge['fromColumn'])) {
                    $sheet->mergeCells("{$merge['fromColumn']}{$merge['onRow']}:{$merge['toColumn']}{$merge['onRow']}");
                } elseif (isset($merge['column'])) {
                    $sheet->mergeCells("{$merge['column']}{$merge['fromRow']}:{$merge['column']}{$merge['toRow']}");
                }
            }
        }

        if (!empty($styles['columnAlignmentStyles'])) {
            foreach ($styles['columnAlignmentStyles'] as $style) {
                $sheet->getStyle($style['range'])->applyFromArray($style['style']);
            }
        }

        if (!empty($styles['cellDataAlign'])) {
            foreach ($styles['cellDataAlign'] as $cellAlign) {
                $sheet->getStyle($cellAlign['range'])->applyFromArray($cellAlign['style']);
            }
        }

        if (!empty($styles['cellBackgroundColor'])) {
            foreach ($styles['cellBackgroundColor'] as $bgColor) {
                $sheet->getStyle($bgColor['range'])->applyFromArray($bgColor['style']);
            }
        }

        if (!empty($styles['rowDataBold'])) {
            foreach ($styles['rowDataBold'] as $rowBold) {
                $sheet->getStyle($rowBold['range'])->applyFromArray($rowBold['style']);
            }
        }

        if (!empty($styles['cellBorderStyle'])) {
            foreach ($styles['cellBorderStyle'] as $border) {
                $sheet->getStyle($border['range'])->applyFromArray($border['style']);
            }
        }
    }

    public function getStyle()
    {
        $mergeCells = [];
        foreach ($this->mainHeaderMergeToCenter ?? [] as $merge) {
            if (isset($merge['fromColumn'], $merge['toColumn'], $merge['onRow'])) {
                $mergeCells[] = $merge;
            }
        }

        foreach ($this->mergeCellToBottom ?? [] as $merge) {
            if (isset($merge['column'], $merge['fromRow'], $merge['toRow'])) {
                $mergeCells[] = $merge;
            }
        }

        foreach ($this->mergeCellToRight ?? [] as $merge) {
            if (isset($merge['fromColumn'], $merge['toColumn'], $merge['onRow'])) {
                $mergeCells[] = $merge;
            }
        }

        $columnAlignmentStyles = [];
        foreach ($this->headerToCenterAlign ?? [] as $item) {
            if (isset($item['fromColumn'], $item['toColumn'], $item['onRow'])) {
                $columnAlignmentStyles[] = [
                    'range' => "{$item['fromColumn']}{$item['onRow']}:{$item['toColumn']}{$item['onRow']}",
                    'style' => $this->getAlignmentStyle('center'),
                ];
            }
        }

        foreach ($this->columnDataToLeftAlign ?? [] as $item) {
            if (isset($item['fromColumn'], $item['toColumn'])) {
                $columnAlignmentStyles[] = [
                    'range' => "{$item['fromColumn']}{$this->startRowData}:{$item['toColumn']}{$this->lastRowData}",
                    'style' => $this->getAlignmentStyle('left'),
                ];
            }
        }

        foreach ($this->columnDataToCenterAlign ?? [] as $item) {
            if (isset($item['fromColumn'], $item['toColumn'])) {
                $columnAlignmentStyles[] = [
                    'range' => "{$item['fromColumn']}{$this->startRowData}:{$item['toColumn']}{$this->lastRowData}",
                    'style' => $this->getAlignmentStyle('center'),
                ];
            }
        }

        foreach ($this->columnDataToRightAlign ?? [] as $item) {
            if (isset($item['fromColumn'], $item['toColumn'])) {
                $columnAlignmentStyles[] = [
                    'range' => "{$item['fromColumn']}{$this->startRowData}:{$item['toColumn']}{$this->lastRowData}",
                    'style' => $this->getAlignmentStyle('right'),
                ];
            }
        }

        $cellDataAlign = [];
        foreach ($this->cellDataAlign ?? [] as $item) {
            if (isset($item['cell'], $item['align'])) {
                $cellDataAlign[] = [
                    'range' => $item['cell'],
                    'style' => $this->getAlignmentStyle($item['align']),
                ];
            }
        }

        $cellBackgroundColor = [];
        foreach ($this->cellBackgroundColor ?? [] as $bgColor) {
            if (isset($bgColor['fromColumn'], $bgColor['toColumn'], $bgColor['onRow'], $bgColor['color'])) {
                $cellBackgroundColor[] = [
                    'range' => "{$bgColor['fromColumn']}{$bgColor['onRow']}:{$bgColor['toColumn']}{$bgColor['onRow']}",
                    'style' => [
                        'fill' => [
                            'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                            'startColor' => ['argb' => $bgColor['color']],
                        ],
                    ],
                ];
            }
        }

        $cellBorderStyle = [];
        foreach ($this->cellBorder ?? [] as $border) {
            if (isset($border['fromColumn'], $border['toColumn'], $border['fromRow'], $border['toRow'], $border['border'])) {
                $cellBorderStyle[] = [
                    'range' => "{$border['fromColumn']}{$border['fromRow']}:{$border['toColumn']}{$border['toRow']}",
                    'style' => $this->getBorderStyle($border['border']),
                ];
            }
        }

        $rowDataBold = [];
        foreach ($this->rowDataBold ?? [] as $item) {
            if (isset($item['fromColumn'], $item['toColumn'], $item['onRow'])) {
                $rowDataBold[] = [
                    'range' => "{$item['fromColumn']}{$item['onRow']}:{$item['toColumn']}{$item['onRow']}",
                    'style' => [
                        'font' => [
                            'bold' => true,
                        ],
                    ],
                ];
            }
        }

        return [
            'mergeCells' => $mergeCells,
            'columnAlignmentStyles' => $columnAlignmentStyles,
            'cellBorderStyle' => $cellBorderStyle,
            'cellDataAlign' => $cellDataAlign,
            'cellBackgroundColor' => $cellBackgroundColor,
            'rowDataBold' => $rowDataBold,
            'lastRow' => $this->lastRowData,
        ];
    }

    public function setMainHeaderMergeToCenter($data)
    {
        $this->mainHeaderMergeToCenter = $data;
    }
    public function setMergeCellToBottom($data)
    {
        $this->mergeCellToBottom = $data;
    }
    public function setMergeCellToRight($data)
    {
        $this->mergeCellToRight = $data;
    }
    public function setHeaderToCenterAlign($data)
    {
        $this->headerToCenterAlign = $data;
    }
    public function setColumnDataToLeftAlign($data)
    {
        $this->columnDataToLeftAlign = $data;
    }
    public function setColumnDataToCenterAlign($data)
    {
        $this->columnDataToCenterAlign = $data;
    }
    public function setColumnDataToRightAlign($data)
    {
        $this->columnDataToRightAlign = $data;
    }

    public function setCellBorder($style)
    {
        $this->cellBorder = $style;
    }

    public function setCellBackgroundColor($color)
    {
        $this->cellBackgroundColor = $color;
    }

    public function setRowDataBold($data)
    {
        $this->rowDataBold = $data;
    }

    public function setCellDataAlign($data)
    {
        $this->cellDataAlign = $data;
    }

    public function getAlignmentStyle($align)
    {
        $alignment = [
            'alignment' => [
                'horizontal' => StyleAlignment::HORIZONTAL_LEFT,
                'vertical' => StyleAlignment::VERTICAL_CENTER,
                'wrapText' => true,
            ],
        ];

        if ($align === 'center') {
            $alignment['alignment']['horizontal'] = StyleAlignment::HORIZONTAL_CENTER;
        } elseif ($align === 'right') {
            $alignment['alignment']['horizontal'] = StyleAlignment::HORIZONTAL_RIGHT;
        }

        return $alignment;
    }

    public function getBorderStyle($borderStyle)
    {
        $defaultBorderStyle = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '000000'],
                ],
            ],
        ];

        if ($borderStyle === 'thick') {
            $defaultBorderStyle['borders']['allBorders']['borderStyle'] = Border::BORDER_THICK;
        } elseif ($borderStyle === 'double') {
            $defaultBorderStyle['borders']['allBorders']['borderStyle'] = Border::BORDER_DOUBLE;
        } elseif ($borderStyle === 'dotted') {
            $defaultBorderStyle['borders']['allBorders']['borderStyle'] = Border::BORDER_DOTTED;
        } elseif ($borderStyle === 'dashed') {
            $defaultBorderStyle['borders']['allBorders']['borderStyle'] = Border::BORDER_DASHED;
        }

        return $defaultBorderStyle;
    }
}

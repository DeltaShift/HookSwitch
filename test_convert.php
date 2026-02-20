<?php

declare(strict_types=1);

require_once __DIR__ . '/file_format_change.php';

$input = $argv[1] ?? (__DIR__ . '/test.csv');
$output = $argv[2] ?? (__DIR__ . '/output.xlsx');
$mode = $argv[3] ?? 'csv_to_xlsx'; // csv_to_xlsx | xlsx_to_csv

if (!in_array($mode, ['csv_to_xlsx', 'xlsx_to_csv'], true)) {
    fwrite(STDERR, "Invalid mode. Use: csv_to_xlsx or xlsx_to_csv\n");
    exit(1);
}

if (!is_file($input)) {
    fwrite(STDERR, "Input file not found: {$input}\n");
    exit(1);
}

function cpuSecondsFromUsage(array $usage): float
{
    $userSec = ($usage['ru_utime.tv_sec'] ?? 0) + (($usage['ru_utime.tv_usec'] ?? 0) / 1_000_000);
    $sysSec = ($usage['ru_stime.tv_sec'] ?? 0) + (($usage['ru_stime.tv_usec'] ?? 0) / 1_000_000);
    return (float) ($userSec + $sysSec);
}

$usageStart = getrusage();
$timeStart = hrtime(true);
$memStart = memory_get_usage(true);

if ($mode === 'xlsx_to_csv') {
    $result = FileFormatChangeService::convertXlsxToCsv($input, $output);
} else {
    $result = FileFormatChangeService::convertCsvToXlsx($input, $output);
}

$timeEnd = hrtime(true);
$usageEnd = getrusage();
$memEnd = memory_get_usage(true);
$peakMem = memory_get_peak_usage(true);

$wallSeconds = ($timeEnd - $timeStart) / 1_000_000_000;
$cpuSeconds = cpuSecondsFromUsage($usageEnd) - cpuSecondsFromUsage($usageStart);
$cpuPercent = $wallSeconds > 0 ? ($cpuSeconds / $wallSeconds) * 100 : 0.0;
$outputSize = is_file($output) ? filesize($output) : 0;

if ($result) {
    echo "SUCCESS: conversion completed\n";
} else {
    echo "FAILED\n";
}

echo "Input: {$input}\n";
echo "Output: {$output}\n";
echo "Mode: {$mode}\n";
echo "Wall time (s): " . number_format($wallSeconds, 6) . "\n";
echo "CPU time (s): " . number_format($cpuSeconds, 6) . "\n";
echo "CPU usage (%): " . number_format($cpuPercent, 2) . "\n";
echo "Memory start (MB): " . number_format($memStart / 1024 / 1024, 2) . "\n";
echo "Memory end (MB): " . number_format($memEnd / 1024 / 1024, 2) . "\n";
echo "Peak memory (MB): " . number_format($peakMem / 1024 / 1024, 2) . "\n";
echo "Output size (bytes): {$outputSize}\n";

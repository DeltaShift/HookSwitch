<?php

declare(strict_types=1);

/**
 * Streaming XLSX/CSV conversion service with low-memory processing.
 */
class FileFormatChangeService
{
    private const PROJECT_ROOT_LEVELS = 3;
    private const OFFSET_RECORD_LENGTH = 21; // 20 digits + newline

    /**
     * Convert an XLSX file (first worksheet) to CSV in a streaming manner.
     */
    public static function convertXlsxToCsv(string $inputPath, string $outputPath): bool
    {
        $input = self::resolveInputPath($inputPath);
        $output = self::resolveOutputPath($outputPath);

        if ($input === null || $output === null) {
            return false;
        }

        $zip = new ZipArchive();
        if ($zip->open($input) !== true) {
            return false;
        }

        $sheetEntry = self::resolveFirstWorksheetEntry($zip);
        if ($sheetEntry === null || $zip->locateName($sheetEntry) === false) {
            $zip->close();
            return false;
        }

        $csvHandle = @fopen($output, 'wb');
        if ($csvHandle === false) {
            $zip->close();
            return false;
        }

        $sheetTempPath = null;
        $sharedStore = null;
        $reader = new XMLReader();
        $success = false;

        try {
            $sheetTempPath = self::copyZipEntryToTempFile($zip, $sheetEntry, 'xlsx_sheet_');
            if ($sheetTempPath === null) {
                return false;
            }

            $sharedStore = self::buildSharedStringStore($zip);

            if (!$reader->open($sheetTempPath, null, LIBXML_NONET | LIBXML_COMPACT | LIBXML_PARSEHUGE)) {
                return false;
            }

            $expectedRowNumber = 1;

            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || self::nodeName($reader) !== 'row') {
                    continue;
                }

                $rowNumberAttr = $reader->getAttribute('r');
                $rowNumber = $rowNumberAttr !== null ? (int) $rowNumberAttr : $expectedRowNumber;
                if ($rowNumber <= 0) {
                    $rowNumber = $expectedRowNumber;
                }

                while ($expectedRowNumber < $rowNumber) {
                    if (@fwrite($csvHandle, "\n") === false) {
                        return false;
                    }
                    $expectedRowNumber++;
                }

                $rowData = self::readRowCells($reader, $sharedStore);
                if ($rowData === null) {
                    return false;
                }

                if (@fputcsv($csvHandle, $rowData) === false) {
                    return false;
                }

                $expectedRowNumber = $rowNumber + 1;
            }

            $success = true;
            return true;
        } catch (Throwable $e) {
            return false;
        } finally {
            $reader->close();

            if (is_resource($csvHandle)) {
                fclose($csvHandle);
            }

            if (is_array($sharedStore)) {
                self::closeSharedStringStore($sharedStore);
            }

            if (is_string($sheetTempPath) && is_file($sheetTempPath)) {
                @unlink($sheetTempPath);
            }

            $zip->close();

            if (!$success && is_file($output)) {
                @unlink($output);
            }
        }
    }

    /**
     * Convert a CSV file to an XLSX file (sheet1) in a streaming manner.
     */
    public static function convertCsvToXlsx(string $inputPath, string $outputPath): bool
    {
        $input = self::resolveInputPath($inputPath);
        $output = self::resolveOutputPath($outputPath);

        if ($input === null || $output === null) {
            return false;
        }

        $csvHandle = @fopen($input, 'rb');
        if ($csvHandle === false) {
            return false;
        }

        $sheetTempPath = tempnam(sys_get_temp_dir(), 'csv_sheet_');
        if ($sheetTempPath === false) {
            fclose($csvHandle);
            return false;
        }

        $sheetHandle = @fopen($sheetTempPath, 'wb');
        if ($sheetHandle === false) {
            fclose($csvHandle);
            @unlink($sheetTempPath);
            return false;
        }

        $zip = new ZipArchive();
        $zipOpened = false;
        $success = false;

        try {
            // Write worksheet XML incrementally while reading CSV row-by-row.
            self::writeString($sheetHandle, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
            self::writeString($sheetHandle, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>');

            $delimiter = self::detectCsvDelimiter($csvHandle);
            if (fseek($csvHandle, 0) !== 0) {
                return false;
            }

            $rowNumber = 0;
            while (($row = fgetcsv($csvHandle, 0, $delimiter, '"', '\\')) !== false) {
                $rowNumber++;
                if ($rowNumber === 1 && isset($row[0])) {
                    $row[0] = self::stripUtf8Bom((string) $row[0]);
                }
                self::writeString($sheetHandle, '<row r="' . $rowNumber . '">');

                $columnCount = count($row);
                for ($column = 0; $column < $columnCount; $column++) {
                    $value = (string) $row[$column];
                    if ($value === '') {
                        continue;
                    }

                    $cellRef = self::columnNumberToName($column + 1) . $rowNumber;
                    $escaped = self::escapeXml(self::sanitizeXmlText($value));
                    self::writeString(
                        $sheetHandle,
                        '<c r="' . $cellRef . '" t="inlineStr"><is><t xml:space="preserve">' . $escaped . '</t></is></c>'
                    );
                }

                self::writeString($sheetHandle, '</row>');
            }

            self::writeString($sheetHandle, '</sheetData></worksheet>');
            fclose($sheetHandle);
            fclose($csvHandle);

            if ($zip->open($output, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
                return false;
            }
            $zipOpened = true;

            // Add minimal XLSX package parts required by Excel-compatible readers.
            if (
                !$zip->addFromString('[Content_Types].xml', self::contentTypesXml()) ||
                !$zip->addFromString('_rels/.rels', self::rootRelsXml()) ||
                !$zip->addFromString('xl/workbook.xml', self::workbookXml()) ||
                !$zip->addFromString('xl/_rels/workbook.xml.rels', self::workbookRelsXml()) ||
                !$zip->addFile($sheetTempPath, 'xl/worksheets/sheet1.xml')
            ) {
                return false;
            }

            if (!$zip->close()) {
                return false;
            }
            $zipOpened = false;

            $success = true;
            return true;
        } catch (Throwable $e) {
            return false;
        } finally {
            if (is_resource($sheetHandle)) {
                fclose($sheetHandle);
            }

            if (is_resource($csvHandle)) {
                fclose($csvHandle);
            }

            if ($zipOpened) {
                $zip->close();
            }

            if (is_file($sheetTempPath)) {
                @unlink($sheetTempPath);
            }

            if (!$success && is_file($output)) {
                @unlink($output);
            }
        }
    }

    /**
     * Resolve the first worksheet entry from workbook relationships.
     */
    private static function resolveFirstWorksheetEntry(ZipArchive $zip): ?string
    {
        $workbookXml = $zip->getFromName('xl/workbook.xml');
        $workbookRelsXml = $zip->getFromName('xl/_rels/workbook.xml.rels');

        if (is_string($workbookXml) && is_string($workbookRelsXml)) {
            $firstRid = self::extractFirstSheetRelationshipId($workbookXml);
            if ($firstRid !== null) {
                $target = self::extractWorksheetTargetByRid($workbookRelsXml, $firstRid);
                if ($target !== null) {
                    $entry = self::normalizeWorksheetTargetToEntry($target);
                    if ($zip->locateName($entry) !== false) {
                        return $entry;
                    }
                }
            }
        }

        $worksheetEntries = [];
        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = $zip->getNameIndex($i);
            if (is_string($name) && preg_match('#^xl/worksheets/[^/]+\.xml$#i', $name) === 1) {
                $worksheetEntries[] = $name;
            }
        }

        if ($worksheetEntries === []) {
            return null;
        }

        sort($worksheetEntries, SORT_STRING);
        return $worksheetEntries[0];
    }

    /**
     * Extract first worksheet relationship id (r:id) from workbook.xml.
     */
    private static function extractFirstSheetRelationshipId(string $workbookXml): ?string
    {
        $reader = new XMLReader();
        if (!$reader->XML($workbookXml, null, LIBXML_NONET | LIBXML_COMPACT | LIBXML_PARSEHUGE)) {
            return null;
        }

        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || self::nodeName($reader) !== 'sheet') {
                    continue;
                }

                $rid = $reader->getAttribute('r:id');
                if ($rid === null || $rid === '') {
                    $rid = $reader->getAttributeNs('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
                }

                if ($rid !== null && $rid !== '') {
                    return $rid;
                }
            }
        } finally {
            $reader->close();
        }

        return null;
    }

    /**
     * Resolve worksheet target by relationship id in workbook.xml.rels.
     */
    private static function extractWorksheetTargetByRid(string $workbookRelsXml, string $rid): ?string
    {
        $reader = new XMLReader();
        if (!$reader->XML($workbookRelsXml, null, LIBXML_NONET | LIBXML_COMPACT | LIBXML_PARSEHUGE)) {
            return null;
        }

        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || self::nodeName($reader) !== 'Relationship') {
                    continue;
                }

                $id = $reader->getAttribute('Id');
                if ($id !== $rid) {
                    continue;
                }

                $target = $reader->getAttribute('Target');
                if ($target === null || $target === '') {
                    return null;
                }

                return $target;
            }
        } finally {
            $reader->close();
        }

        return null;
    }

    /**
     * Normalize worksheet target path into ZIP entry path.
     */
    private static function normalizeWorksheetTargetToEntry(string $target): string
    {
        $normalized = str_replace('\\', '/', $target);
        if (str_starts_with($normalized, '/')) {
            return ltrim($normalized, '/');
        }
        if (str_starts_with($normalized, 'xl/')) {
            return $normalized;
        }
        return 'xl/' . ltrim($normalized, '/');
    }

    /**
     * Read one row from the current <row> element using XMLReader streaming.
     *
     * @param array<string, mixed>|null $sharedStore
     * @return array<int, string>|null
     */
    private static function readRowCells(XMLReader $reader, ?array $sharedStore): ?array
    {
        $rowDepth = $reader->depth;
        $rowMap = [];
        $maxColumn = 0;
        $currentColumn = 1;

        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::END_ELEMENT && self::nodeName($reader) === 'row' && $reader->depth === $rowDepth) {
                break;
            }

            if ($reader->nodeType !== XMLReader::ELEMENT || self::nodeName($reader) !== 'c') {
                continue;
            }

            $cellRef = $reader->getAttribute('r');
            $cellType = $reader->getAttribute('t') ?? '';

            $columnIndex = $cellRef !== null
                ? self::columnNameToNumber(self::extractColumnLetters($cellRef))
                : $currentColumn;

            if ($columnIndex <= 0) {
                $columnIndex = $currentColumn;
            }

            $rawValue = self::readCellValue($reader);
            if ($rawValue === null) {
                return null;
            }

            if ($cellType === 's') {
                $sharedIndex = (int) $rawValue;
                $value = $sharedStore !== null ? self::readSharedStringByIndex($sharedStore, $sharedIndex) : '';
                if ($value === null) {
                    return null;
                }
            } elseif ($cellType === 'b') {
                $value = ((string) $rawValue === '1') ? 'TRUE' : 'FALSE';
            } else {
                $value = $rawValue;
            }

            $rowMap[$columnIndex] = $value;
            if ($columnIndex > $maxColumn) {
                $maxColumn = $columnIndex;
            }

            $currentColumn = $columnIndex + 1;
        }

        if ($maxColumn === 0) {
            return [];
        }

        $row = [];
        for ($column = 1; $column <= $maxColumn; $column++) {
            $row[] = $rowMap[$column] ?? '';
        }

        return $row;
    }

    /**
     * Read one cell value from the current <c> element.
     */
    private static function readCellValue(XMLReader $reader): ?string
    {
        if ($reader->isEmptyElement) {
            return '';
        }

        $cellDepth = $reader->depth;
        $value = '';

        while ($reader->read()) {
            if ($reader->nodeType === XMLReader::END_ELEMENT && self::nodeName($reader) === 'c' && $reader->depth === $cellDepth) {
                break;
            }

            if ($reader->nodeType !== XMLReader::ELEMENT) {
                continue;
            }

            if (self::nodeName($reader) === 'v') {
                $value = $reader->readString();
            } elseif (self::nodeName($reader) === 't') {
                // Handles inlineStr and rich text nodes.
                $value .= $reader->readString();
            }
        }

        return $value;
    }

    /**
     * Build a disk-backed shared string store with O(1)-style offset lookups.
     *
     * @return array<string, mixed>|null
     */
    private static function buildSharedStringStore(ZipArchive $zip): ?array
    {
        $entry = 'xl/sharedStrings.xml';
        if ($zip->locateName($entry) === false) {
            return null;
        }

        $xmlTempPath = self::copyZipEntryToTempFile($zip, $entry, 'xlsx_sst_xml_');
        if ($xmlTempPath === null) {
            return null;
        }

        $dataPath = tempnam(sys_get_temp_dir(), 'xlsx_sst_dat_');
        $indexPath = tempnam(sys_get_temp_dir(), 'xlsx_sst_idx_');
        if ($dataPath === false || $indexPath === false) {
            @unlink($xmlTempPath);
            if (is_string($dataPath)) {
                @unlink($dataPath);
            }
            if (is_string($indexPath)) {
                @unlink($indexPath);
            }
            return null;
        }

        $dataHandle = @fopen($dataPath, 'wb+');
        $indexHandle = @fopen($indexPath, 'wb+');
        if ($dataHandle === false || $indexHandle === false) {
            @unlink($xmlTempPath);
            @unlink($dataPath);
            @unlink($indexPath);
            if (is_resource($dataHandle)) {
                fclose($dataHandle);
            }
            if (is_resource($indexHandle)) {
                fclose($indexHandle);
            }
            return null;
        }

        $reader = new XMLReader();
        if (!$reader->open($xmlTempPath, null, LIBXML_NONET | LIBXML_COMPACT | LIBXML_PARSEHUGE)) {
            fclose($dataHandle);
            fclose($indexHandle);
            @unlink($xmlTempPath);
            @unlink($dataPath);
            @unlink($indexPath);
            return null;
        }

        try {
            while ($reader->read()) {
                if ($reader->nodeType !== XMLReader::ELEMENT || self::nodeName($reader) !== 'si') {
                    continue;
                }

                $siXml = $reader->readOuterXML();
                if ($siXml === '') {
                    return null;
                }

                $text = self::extractSharedStringText($siXml);
                if ($text === null) {
                    return null;
                }

                $offset = ftell($dataHandle);
                if ($offset === false) {
                    return null;
                }

                // Fixed-width offset records allow direct seek by shared string index.
                if (@fwrite($indexHandle, sprintf('%020d\n', $offset)) !== self::OFFSET_RECORD_LENGTH) {
                    return null;
                }

                $len = strlen($text);
                if (@fwrite($dataHandle, pack('N', $len)) !== 4) {
                    return null;
                }
                if ($len > 0 && @fwrite($dataHandle, $text) !== $len) {
                    return null;
                }
            }

            fflush($dataHandle);
            fflush($indexHandle);

            @unlink($xmlTempPath);

            return [
                'data_path' => $dataPath,
                'index_path' => $indexPath,
                'data_handle' => $dataHandle,
                'index_handle' => $indexHandle,
            ];
        } catch (Throwable $e) {
            fclose($dataHandle);
            fclose($indexHandle);
            @unlink($xmlTempPath);
            @unlink($dataPath);
            @unlink($indexPath);
            return null;
        } finally {
            $reader->close();
        }
    }

    /**
     * Read a shared string by index from a disk-backed store.
     */
    private static function readSharedStringByIndex(array $store, int $index): ?string
    {
        if ($index < 0) {
            return '';
        }

        $offsetHandle = $store['index_handle'] ?? null;
        $dataHandle = $store['data_handle'] ?? null;
        if (!is_resource($offsetHandle) || !is_resource($dataHandle)) {
            return null;
        }

        $offsetPos = $index * self::OFFSET_RECORD_LENGTH;
        if (fseek($offsetHandle, $offsetPos, SEEK_SET) !== 0) {
            return '';
        }

        $record = fread($offsetHandle, self::OFFSET_RECORD_LENGTH);
        if (!is_string($record) || strlen($record) !== self::OFFSET_RECORD_LENGTH) {
            return '';
        }

        $offset = (int) trim($record);
        if ($offset < 0) {
            return '';
        }

        if (fseek($dataHandle, $offset, SEEK_SET) !== 0) {
            return '';
        }

        $lenBytes = fread($dataHandle, 4);
        if (!is_string($lenBytes) || strlen($lenBytes) !== 4) {
            return '';
        }

        $len = unpack('Nlen', $lenBytes);
        $size = (int) ($len['len'] ?? 0);
        if ($size <= 0) {
            return '';
        }

        $data = '';
        while (strlen($data) < $size) {
            $chunk = fread($dataHandle, $size - strlen($data));
            if ($chunk === false || $chunk === '') {
                return '';
            }
            $data .= $chunk;
        }

        return $data;
    }

    /**
     * Close and remove temporary shared string files.
     */
    private static function closeSharedStringStore(array $store): void
    {
        $dataHandle = $store['data_handle'] ?? null;
        $indexHandle = $store['index_handle'] ?? null;
        $dataPath = $store['data_path'] ?? null;
        $indexPath = $store['index_path'] ?? null;

        if (is_resource($dataHandle)) {
            fclose($dataHandle);
        }
        if (is_resource($indexHandle)) {
            fclose($indexHandle);
        }
        if (is_string($dataPath) && is_file($dataPath)) {
            @unlink($dataPath);
        }
        if (is_string($indexPath) && is_file($indexPath)) {
            @unlink($indexPath);
        }
    }

    /**
     * Extract shared string text from one <si> XML chunk.
     */
    private static function extractSharedStringText(string $siXml): ?string
    {
        $reader = new XMLReader();
        if (!$reader->XML($siXml, null, LIBXML_NONET | LIBXML_COMPACT | LIBXML_PARSEHUGE)) {
            return null;
        }

        $text = '';
        try {
            while ($reader->read()) {
                if ($reader->nodeType === XMLReader::ELEMENT && self::nodeName($reader) === 't') {
                    $text .= $reader->readString();
                }
            }
            return $text;
        } finally {
            $reader->close();
        }
    }

    /**
     * Read XMLReader node name in a namespace-safe way.
     */
    private static function nodeName(XMLReader $reader): string
    {
        return $reader->localName !== '' ? $reader->localName : $reader->name;
    }

    /**
     * Detect CSV delimiter from first non-empty line.
     */
    private static function detectCsvDelimiter($csvHandle): string
    {
        $delimiters = [',', ';', "\t", '|'];
        $line = '';

        while (!feof($csvHandle)) {
            $candidate = fgets($csvHandle);
            if ($candidate === false) {
                break;
            }

            if (trim($candidate) === '') {
                continue;
            }

            $line = $candidate;
            break;
        }

        if ($line === '') {
            return ',';
        }

        $line = self::stripUtf8Bom($line);
        $bestDelimiter = ',';
        $bestCount = -1;

        foreach ($delimiters as $delimiter) {
            $fields = str_getcsv($line, $delimiter, '"', '\\');
            $count = count($fields);
            if ($count > $bestCount) {
                $bestCount = $count;
                $bestDelimiter = $delimiter;
            }
        }

        return $bestDelimiter;
    }

    /**
     * Remove UTF-8 BOM from start of a string.
     */
    private static function stripUtf8Bom(string $value): string
    {
        if (strncmp($value, "\xEF\xBB\xBF", 3) === 0) {
            return substr($value, 3);
        }
        return $value;
    }

    /**
     * Copy one ZIP entry to a temporary file without loading full content in memory.
     */
    private static function copyZipEntryToTempFile(ZipArchive $zip, string $entryName, string $prefix): ?string
    {
        $source = $zip->getStream($entryName);
        if ($source === false) {
            return null;
        }

        $tempPath = tempnam(sys_get_temp_dir(), $prefix);
        if ($tempPath === false) {
            fclose($source);
            return null;
        }

        $dest = @fopen($tempPath, 'wb');
        if ($dest === false) {
            fclose($source);
            @unlink($tempPath);
            return null;
        }

        $ok = true;
        while (!feof($source)) {
            $chunk = fread($source, 1024 * 1024);
            if ($chunk === false) {
                $ok = false;
                break;
            }
            if ($chunk !== '' && @fwrite($dest, $chunk) === false) {
                $ok = false;
                break;
            }
        }

        fclose($source);
        fclose($dest);

        if (!$ok) {
            @unlink($tempPath);
            return null;
        }

        return $tempPath;
    }

    /**
     * Validate and resolve an existing input file path.
     */
    private static function resolveInputPath(string $path): ?string
    {
        if (!self::isSafePathFormat($path)) {
            return null;
        }

        $resolved = realpath($path);
        if ($resolved === false || !is_file($resolved) || !is_readable($resolved)) {
            return null;
        }

        return self::isAllowedResolvedPath($resolved) ? $resolved : null;
    }

    /**
     * Validate and resolve an output file path.
     */
    private static function resolveOutputPath(string $path): ?string
    {
        if (!self::isSafePathFormat($path)) {
            return null;
        }

        $directory = realpath(dirname($path));
        if ($directory === false || !is_dir($directory) || !is_writable($directory)) {
            return null;
        }

        if (!self::isAllowedResolvedPath($directory)) {
            return null;
        }

        return $directory . DIRECTORY_SEPARATOR . basename($path);
    }

    /**
     * Basic path safety checks against traversal wrappers/injections.
     */
    private static function isSafePathFormat(string $path): bool
    {
        if ($path === '' || str_contains($path, "\0")) {
            return false;
        }

        if (preg_match('/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//', $path) === 1) {
            return false;
        }

        $normalized = str_replace('\\\\', '/', $path);
        foreach (explode('/', $normalized) as $part) {
            if ($part === '..') {
                return false;
            }
        }

        return true;
    }

    /**
     * Ensure resolved path remains inside the project root.
     */
    private static function isUnderProjectRoot(string $resolvedPath): bool
    {
        $root = realpath(__DIR__ . str_repeat('/..', self::PROJECT_ROOT_LEVELS));
        if ($root === false) {
            return false;
        }

        $normalizedRoot = rtrim(str_replace('\\\\', '/', $root), '/') . '/';
        $normalizedPath = str_replace('\\\\', '/', $resolvedPath);

        return str_starts_with($normalizedPath . '/', $normalizedRoot) || $normalizedPath === rtrim($normalizedRoot, '/');
    }

    /**
     * Allow access only inside project root or system temp directory.
     */
    private static function isAllowedResolvedPath(string $resolvedPath): bool
    {
        if (self::isUnderProjectRoot($resolvedPath)) {
            return true;
        }

        $tempDir = realpath(sys_get_temp_dir());
        if ($tempDir === false) {
            return false;
        }

        $normalizedTemp = rtrim(str_replace('\\', '/', $tempDir), '/') . '/';
        $normalizedPath = str_replace('\\', '/', $resolvedPath);

        return str_starts_with($normalizedPath . '/', $normalizedTemp) || $normalizedPath === rtrim($normalizedTemp, '/');
    }

    /**
     * Convert cell reference column letters (e.g. "BC") to a 1-based number.
     */
    private static function columnNameToNumber(string $letters): int
    {
        $letters = strtoupper($letters);
        $length = strlen($letters);
        $number = 0;

        for ($i = 0; $i < $length; $i++) {
            $char = ord($letters[$i]);
            if ($char < 65 || $char > 90) {
                continue;
            }
            $number = ($number * 26) + ($char - 64);
        }

        return $number;
    }

    /**
     * Convert a 1-based column number to Excel letters.
     */
    private static function columnNumberToName(int $column): string
    {
        if ($column < 1) {
            return 'A';
        }

        $name = '';
        while ($column > 0) {
            $column--;
            $name = chr(($column % 26) + 65) . $name;
            $column = intdiv($column, 26);
        }

        return $name;
    }

    /**
     * Extract the alphabetical column segment from a cell reference.
     */
    private static function extractColumnLetters(string $cellRef): string
    {
        if (preg_match('/^[A-Z]+/i', $cellRef, $match) === 1) {
            return strtoupper($match[0]);
        }

        return 'A';
    }

    /**
     * XML escape helper for worksheet text nodes.
     */
    private static function escapeXml(string $value): string
    {
        return htmlspecialchars($value, ENT_XML1 | ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
    }

    /**
     * Remove bytes/chars that are illegal in XML 1.0 text nodes.
     */
    private static function sanitizeXmlText(string $value): string
    {
        // Keep valid XML chars only: TAB, LF, CR, and legal Unicode scalar ranges.
        $clean = preg_replace('/[^\x09\x0A\x0D\x20-\x{D7FF}\x{E000}-\x{FFFD}\x{10000}-\x{10FFFF}]/u', '', $value);
        if ($clean !== null) {
            return $clean;
        }

        // Fallback when input contains invalid UTF-8 byte sequences.
        if (function_exists('iconv')) {
            $utf8 = @iconv('UTF-8', 'UTF-8//IGNORE', $value);
            if (is_string($utf8)) {
                $retry = preg_replace('/[^\x09\x0A\x0D\x20-\x{D7FF}\x{E000}-\x{FFFD}\x{10000}-\x{10FFFF}]/u', '', $utf8);
                if ($retry !== null) {
                    return $retry;
                }
                return $utf8;
            }
        }

        return $value;
    }

    /**
     * Safe fwrite wrapper that throws on write failures.
     */
    private static function writeString($handle, string $value): void
    {
        $length = strlen($value);
        $written = 0;

        while ($written < $length) {
            $result = fwrite($handle, substr($value, $written));
            if ($result === false || $result === 0) {
                throw new RuntimeException('Write failed while creating XLSX stream.');
            }
            $written += $result;
        }
    }

    /**
     * Static XML payload for [Content_Types].xml.
     */
    private static function contentTypesXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            . '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            . '<Default Extension="xml" ContentType="application/xml"/>'
            . '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            . '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            . '</Types>';
    }

    /**
     * Static XML payload for _rels/.rels.
     */
    private static function rootRelsXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            . '</Relationships>';
    }

    /**
     * Static XML payload for xl/workbook.xml.
     */
    private static function workbookXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            . 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            . '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>'
            . '</workbook>';
    }

    /**
     * Static XML payload for xl/_rels/workbook.xml.rels.
     */
    private static function workbookRelsXml(): string
    {
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            . '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            . '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            . '</Relationships>';
    }
}

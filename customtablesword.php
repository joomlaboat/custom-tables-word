<?php
/**
 * CustomTablesWord Joomla! 3.x Native Component
 * @author Ivan Komlev <support@joomlaboat.com>
 * @link http://www.joomlaboat.com
 * @copyright (C) 2018-2024 Ivan Komlev
 * @license GNU/GPL
 **/

namespace CustomTables;

// no direct access
if (!defined('_JEXEC') and !defined('WPINC')) {
    die('Restricted access');
}

use CustomTables\CTMiscHelper;
use DOMDocument;
use PhpOffice\PhpWord\Element\Table;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\SimpleType\TblWidth;

class Twig_PHPWord_Tags
{
    function setvalues(string $templateFile, array $values, ?string $saveAsFile = null)
    {
        //Generates a new .docx file and returns the link to that file if successful.

        if ($templateFile == '') {
            echo 'Template file name not specified.';
            return null;
        }

        if ($templateFile[0] != '/' and $templateFile[0] != '\\')
            $templateFile = JPATH_SITE . DIRECTORY_SEPARATOR . $templateFile;

        if (!file_exists($templateFile)) {
            echo 'File "' . $templateFile . '" not exists.';
            return null;
        }

        $webFileLink = '';
        if ($saveAsFile === null) {
            $saveAsFile = CTMiscHelper::suggest_TempFileName($webFileLink, 'docx');
        } else {
            if (file_exists($saveAsFile)) {
                echo 'File "' . $saveAsFile . '" already exists.';
                return null;
            }
        }

        require_once 'bootstrap.php';

        $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor($templateFile);

        foreach ($values as $key => $item) {

            if (str_contains(strtolower($item), '<table')) {
                $item = mb_convert_encoding($item, 'UTF-8', 'UTF-8');
                $table = $this->parseHTMLTable($item);
                $templateProcessor->setComplexBlock('Table', $table);
            } else {
                $item = strip_tags($item);
                $utf8_string = mb_convert_encoding($item, 'UTF-8', 'UTF-8');
                $templateProcessor->setValue($key, $utf8_string);
            }
        }

        $templateProcessor->saveAs($saveAsFile);
        return $webFileLink;
    }

    protected function parseHTMLTable($string)
    {
        $dom = new domDocument;
        @$dom->loadHTML($string);
        $dom->preserveWhiteSpace = false;
        $tables = $dom->getElementsByTagName('table');

        $rows = $tables->item(0)->getElementsByTagName('tr');
        if (count($rows) == 0)
            return null;

        $document_with_table = new PhpWord();
        //$section = $document_with_table->addSection();
        //$table = $section->addTable();//'myOwnTableStyle');
        $table = new Table(array('borderSize' => 7, 'borderColor' => '000000', 'width' => 8500, 'unit' => TblWidth::TWIP));

        //$styleCell = array('borderTopSize'=>1 ,'borderTopColor' =>'black','borderLeftSize'=>1,'borderLeftColor' =>'black','borderRightSize'=>1,'borderRightColor'=>'black','borderBottomSize' =>1,'borderBottomColor'=>'black' );
        $TfontStyle = array('bold' => false, 'italic' => false, 'size' => 12, 'name' => 'Times New Roman', 'afterSpacing' => 0, 'Spacing' => 0, 'cellMargin' => 0);

        foreach ($rows as $row) {

            $cols = $row->getElementsByTagName('th');
            if (count($cols) > 0) {
                $table->addRow();
                $count = 0;
                foreach ($cols as $col) {
                    $count += 1;
                    $item = utf8_decode($col->textContent);//mb_convert_encoding($col->textContent, 'UTF-7','UTF-8');
                    if ($count == 1)
                        $table->addCell(700)->addText($item, $TfontStyle);
                    else
                        $table->addCell(2500)->addText($item, $TfontStyle);
                }
            }

            $cols = $row->getElementsByTagName('td');
            if (count($cols) > 0) {
                $table->addRow();
                $count = 0;
                foreach ($cols as $col) {
                    $count += 1;
                    $item = utf8_decode($col->textContent);//mb_convert_encoding($col->textContent, 'UTF-7','UTF-8');
                    if ($count == 1)
                        $table->addCell(700)->addText($item, $TfontStyle);
                    else
                        $table->addCell(2500)->addText($item, $TfontStyle);
                }
            }
        }
        return $table;
    }
}
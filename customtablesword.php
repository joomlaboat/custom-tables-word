<?php
/**
 * CustomTablesWord Joomla! 3.x Native Component
 * @author Ivan Komlev <support@joomlaboat.com>
 * @link http://www.joomlaboat.com
 * @copyright (C) 2018-2023 Ivan Komlev
 * @license GNU/GPL
 **/

namespace CustomTables;

// no direct access
if (!defined('_JEXEC') and !defined('WPINC')) {
    die('Restricted access');
}

use JoomlaBasicMisc;

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
            $saveAsFile = JoomlaBasicMisc::suggest_TempFileName($webFileLink,'docx');
        } else {
            if (file_exists($saveAsFile)) {
                echo 'File "' . $saveAsFile . '" already exists.';
                return null;
            }
        }

        require_once 'bootstrap.php';

        $templateProcessor = new \PhpOffice\PhpWord\TemplateProcessor($templateFile);

        foreach ($values as $key => $item) {
            $templateProcessor->setValue($key, $item);
        }

        $templateProcessor->saveAs($saveAsFile);
        return $webFileLink;
    }
}
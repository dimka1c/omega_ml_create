<?php

    function deleteAllFilesFromDirectory($path) { //удаление всех файлов из папки $path
        if (file_exists($path)) {
            foreach (glob($path .'/*.xls') as $file) {
                unlink($file);
            }
        }
    }


    function delPathFromNameFile($filename)
    {
        $len = strlen(DIR_ATTACH_SAVE . '/');
        //echo '<br>' . $len;
        foreach($filename as $name) {
            $d = substr($name, $len);
            $arr[] = $d;
        }
        return $arr;
    }


    function isFileName($fname) {
        $f = substr($fname, 0, 2);
        if($f == 'ml') {
            return 1;
        }
        return 0;
    }

    function getBooksDate($subj) {
        $f = substr($subj, 3, 10);
        return $f;
    }


    function chekFilesInFolder($files, $path) {
        if (file_exists($path)) {
            foreach ($files as $file) {
                if(file_exists($path . '/' . $file)) {
                } else {
                    return false;
                }
            }
        } else {
            return false;
        }
        return true;
    }

    function getMlFileName($name) {
         //   ml/07_12_2016/create
        $name = explode('/', $name);
        if(isset($name[1])) {
            return 'ml_' . $name[1] . '.xls';
        } else {
            $today = date("d_m_y");
            return 'ml_' . $today . '.xls';
        }
    }

    function rus2translit($string) {
        $converter = array(
            'а' => 'a',   'б' => 'b',   'в' => 'v',
            'г' => 'g',   'д' => 'd',   'е' => 'e',
            'ё' => 'e',   'ж' => 'zh',  'з' => 'z',
            'и' => 'i',   'й' => 'y',   'к' => 'k',
            'л' => 'l',   'м' => 'm',   'н' => 'n',
            'о' => 'o',   'п' => 'p',   'р' => 'r',
            'с' => 's',   'т' => 't',   'у' => 'u',
            'ф' => 'f',   'х' => 'h',   'ц' => 'c',
            'ч' => 'ch',  'ш' => 'sh',  'щ' => 'sch',
            'ь' => '\'',  'ы' => 'y',   'ъ' => '\'',
            'э' => 'e',   'ю' => 'yu',  'я' => 'ya',

            'А' => 'A',   'Б' => 'B',   'В' => 'V',
            'Г' => 'G',   'Д' => 'D',   'Е' => 'E',
            'Ё' => 'E',   'Ж' => 'Zh',  'З' => 'Z',
            'И' => 'I',   'Й' => 'Y',   'К' => 'K',
            'Л' => 'L',   'М' => 'M',   'Н' => 'N',
            'О' => 'O',   'П' => 'P',   'Р' => 'R',
            'С' => 'S',   'Т' => 'T',   'У' => 'U',
            'Ф' => 'F',   'Х' => 'H',   'Ц' => 'C',
            'Ч' => 'Ch',  'Ш' => 'Sh',  'Щ' => 'Sch',
            'Ь' => '\'',  'Ы' => 'Y',   'Ъ' => '\'',
            'Э' => 'E',   'Ю' => 'Yu',  'Я' => 'Ya',
        );
        return strtr($string, $converter);
    }

    //транслитерация
    function str2url($str) {
        // переводим в транслит
        $str = rus2translit($str);
        // в нижний регистр
        $str = strtolower($str);
        // заменям все ненужное нам на "-"
        //$str = preg_replace('~[^-a-z0-9_]+~u', '-', $str);
        // удаляем начальные и конечные '-'
        //$str = trim($str, "-");
        return $str;
    }


    function debug($data) {
	echo '<pre>';
	var_dump($data);
	echo '</pre>';
	echo '<hr>';
    }


?>
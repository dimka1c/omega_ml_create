<?php

class workExcel
{

    private $objReader;
    private $objWriter;
    //private $pExcel;
    //private $objPHPExcel;
    //private $rowIterator;
    private $cellIterator;
    private $csvWirter;
    private $objWorksheet;
    private $file1;
    private $new_sheet;
    private $sheet;


    // удляет путь к файлу
    //  files/ml0149117_8511.xls -->> ml0149117_8511.xls
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

    // преобразуем 'Водитель:  Лозовский Андрей Сергеевич' в 'Лозовский'
    private function readName($str) {
        $name =  trim(stristr($str, ' '));
        $name1 = substr($name,0,strpos(trim($name),' '));
        return strtolower($name1);
    }

    // преобразуем строку из ячейки AH1 'ШинаКрРогПТ, 51КрРогПТ' в '51'
    private function readDay($str) {
        foreach (WORK_DAY as $day) {
            $n = stripos($str, $day);
            if($n !== false) {
                $d = substr($str, $n, 2);
                return $d;
            }
        }
        return 0;
    }


    //***********************************************************
    //************* редактируем все аттач-файлы и сохраняем *****
    //************* путь для сохранения config.php/ DIR_SAVE
    // возвращаем массив с данными в формате:
    // день выгрузки, название файла ml
    //***********************************************************
    public function editAttachFiles($attach, $type) {

//include 'class/PHPExcel.php';
        if($type == 'ml') {
            $arr_edit_attach = Array();
            //уаляем все файлы из папки edit
            deleteAllFilesFromDirectory(DIR_SAVE_EDIT_ATTACH);
            $kol = 0; //колво файлов
            $all_files = glob(DIR_ATTACH_SAVE .'/*.xls');
            $files = delPathFromNameFile($all_files);
//PHPExcel_Settings::setCacheStorageMethod(PHPExcel_CachedObjectStorageFactory::cache_to_sqlite);

            $pExcel = new PHPExcel();
            $objReader = PHPExcel_IOFactory::createReader('Excel5');

            foreach ($files as $filename) {
                if(strripos($filename, 'ReestrVozvratov') === FALSE) { //не реестр возвратов
                    $objPHPExcel = $objReader->load(DIR_ATTACH_SAVE . '/' . $filename);
                    $objWorksheet = $objPHPExcel->getActiveSheet();

                    // тут должны получить день выгрузки ('51', '12', '23', '34', '45', '56')
                    // находится в ячейке AH1
                    // может иметь такие значения в ячейке?
                    //  - 45ДнепрЧТ, ШинаДнрЧТ
                    //  - 45
                    $day = $objWorksheet->getCell('AH1')->getValue();
                    $day = $this->readDay($day);
                    if($day == 0) { //значение дня неопределено, надо назначить день
                        //может быть общий файл, иногда попадает в выборку, тогда значение надо искать
                        // в ячейке V1
                        $day = $objWorksheet->getCell('V1')->getValue();
                        $day = $this->readDay($day);
                        if($day == 0) {
                            $day = 'NA'; //значение для неопределнныъ дней
                        }
                    }
                    $arr_edit_attach[$kol]['filename'] = $filename;
                    $arr_edit_attach[$kol]['day'] = $day;
                    /*
                     *  Удаление ненужных строк
                     */

                    $marshrut_list = $objWorksheet->getCell('B1')->getValue();
                    $phoneDriver = $objWorksheet->getCell('F5')->getValue();
                    $reisDriver = $objWorksheet->getCell('F4')->getValue();

                    $row_id = 6; // deleted row id
                    $number_rows = 4; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeRow($row_id, $number_rows)) {
                        }
                    }

                    /*
                     *  Удаляем ненужные столбцы
                     */
                    $column_id = 'Q'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'L'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'I'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'F'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'B'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }

                    $objWorksheet->setCellValue('A1', $marshrut_list); // вставляем запомненный ранее маршрутный лист
                    $objWorksheet->setCellValue('D4', $reisDriver); // вставляем запомненный ранее номер рейса
                    $objWorksheet->setCellValue('D5', $phoneDriver); // вставляем запомненный ранее телефон водителя

                    $style_header = array(
                        'font'  => array(
                            'bold'  => true,
                            //'color' => array('rgb' => '778899'),
                            'size'  => 16,
                            'name'  => 'Arial Black'
                        ));
                    //$last_column = $objWorksheet->getHighestColumn()-1;
                    $objWorksheet->getStyle('A1')->applyFromArray($style_header);
                    $objWorksheet->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                    $objWorksheet->getPageSetup()->SetPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                    $objWorksheet->getPageMargins()->setTop(0.1);
                    $objWorksheet->getPageMargins()->setRight(0.1);
                    $objWorksheet->getPageMargins()->setLeft(0.1);
                    $objWorksheet->getPageMargins()->setBottom(0.1);
                    $style_table = array(
                        'font'  => array(
                            'size'  => 10,
                        ));
                    $objWorksheet->getStyle('A7:M100')->applyFromArray($style_table);
                    $style_header_table = array(
                        // Заполнение цветом
                        'fill' => array(
                            'type' => PHPExcel_STYLE_FILL::FILL_SOLID,
                            'color'=>array(
                                'rgb' => 'FFFFFF'
                            )
                        ),
                    );
                    $objWorksheet->getStyle('A1:M6')->applyFromArray($style_header_table);
	
                    /*
                                        $row_id = 3; // deleted row id
                                        $number_rows = 2; // number of rows count
                                        if ($objWorksheet != NULL) {
                                            if ($objWorksheet->removeRow($row_id, $number_rows)) {
                                            }
                                        }
                                        $row_id = 5; // deleted row id
                                        $number_rows = 4; // number of rows count
                                        if ($objWorksheet != NULL) {
                                            if ($objWorksheet->removeRow($row_id, $number_rows)) {
                                            }
                                        }
                    */

                    //цикл по строкам
/*
                    $i = 0;
                    $flag = true;
                    $rowIterator = $objWorksheet->getRowIterator();
                    foreach ($rowIterator as $row) {
                        if ($flag == false) {
                            break;
                        }
                        // Получили ячейки текущей строки и обойдем их в цикле
                        $i++;
                        $cellIterator = $row->getCellIterator();
                        //print_r($cellIterator);
                        foreach ($cellIterator as $cell) {
                            //echo '<br> <b></b>cell = ' . $cell . '</b><br>';
                            if (stripos($cell->getCalculatedValue(), 'Итоговая стоимость') !== FALSE) { //нашли последнюю строку
                                $flag = false;
                                $last_column = $cell->getColumn();
                                $last_row = $cell->getRow();
                                break;
                            }
                        }
                    }
                    $row_id = $i + 1; // deleted row id
                    $number_rows = 100; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeRow($row_id, $number_rows)) {
                        }
                    }
*/
                    // ************************
                    //удаляем ненужные столбцы
                    // ************************
/*
                    $column_id = 'D'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'E'; // deleted row id
                    $number_columns = 1; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'G'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'H'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'I'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $column_id = 'M'; // deleted row id
                    $number_columns = 1; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }

                    $column_id = 'N'; // deleted row id
                    $number_columns = 2; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
*/

/*
                    $column_id = 'W'; // deleted row id
                    $number_columns = 1; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
*/

/*
                    $column_id = 'U'; // deleted row id
                    $number_columns = 1; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }

                    $column_id = 'V'; // deleted row id
                    $number_columns = 1; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }

                    $cell_B1 = $objWorksheet->getCell('B1')->getValue();
                    $cell_B2 = $objWorksheet->getCell('B2')->getValue();
                    $column_id = 'B'; // deleted row id
                    $number_columns = 1; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $cell_C3 = $objWorksheet->getCell('C3')->getValue();
                    $cell_C4 = $objWorksheet->getCell('C4')->getValue();
                    $column_id = 'C'; // deleted row id
                    $number_columns = 1; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeColumn($column_id, $number_columns)) {
                        }
                    }
                    $objWorksheet->setCellValue('E1', $cell_B1);

                    //$objWorksheet->setCellValue('E2', $cell_B2); // маршрут
                    // объединяем ячейки D2-J2
                    // и уже вставляем не в E2, а в D2
                    $objWorksheet->mergeCells('A2:S2');
                    $objWorksheet->setCellValue('A2', $cell_B2); // маршрут

                    $objWorksheet->setCellValue('E3', $cell_C3);
                    $objWorksheet->setCellValue('E4', $cell_C4);

                    $cell_E3 = $objWorksheet->getCell('E3')->getValue();
                    $objWorksheet->setCellValue('E3', '');
                    $objWorksheet->setCellValue('K4', $cell_E3);

                    // настройка параметров страницы
                    //------------------------------------------------------------------
                    $objWorksheet->getPageSetup()
                        ->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
                    $objWorksheet->getPageSetup()
                        ->SetPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
                    $objWorksheet->getPageMargins()->setTop(0.1);
                    $objWorksheet->getPageMargins()->setRight(0.1);
                    $objWorksheet->getPageMargins()->setLeft(0.1);
                    $objWorksheet->getPageMargins()->setBottom(0.1);
                    // задаем ширину столбцов
                    $objWorksheet->getColumnDimension('A')->setWidth(3.5);
                    $objWorksheet->getColumnDimension('B')->setWidth(23);
                    $objWorksheet->getColumnDimension('C')->setWidth(10);
                    $objWorksheet->getColumnDimension('D')->setWidth(27);
                    $objWorksheet->getColumnDimension('E')->setWidth(17);
                    $objWorksheet->getColumnDimension('F')->setWidth(14);
                    $objWorksheet->getColumnDimension('G')->setWidth(4);
                    $objWorksheet->getColumnDimension('H')->setWidth(4);
                    $objWorksheet->getColumnDimension('I')->setWidth(4);
                    $objWorksheet->getColumnDimension('J')->setWidth(6);
                    $objWorksheet->getColumnDimension('K')->setWidth(18);
                    $objWorksheet->getColumnDimension('L')->setWidth(5);
                    $objWorksheet->getColumnDimension('M')->setWidth(4);
                    $objWorksheet->getColumnDimension('N')->setWidth(9);
                    $objWorksheet->getColumnDimension('O')->setWidth(6);
                    $objWorksheet->getColumnDimension('P')->setWidth(7);
                    $objWorksheet->getColumnDimension('Q')->setWidth(6);
                    $objWorksheet->getColumnDimension('R')->setWidth(6);
                    $objWorksheet->getColumnDimension('S')->setWidth(9);
                    //--------------------------------------------------------------
                    //echo $objWorksheet->getHighestColumn() . '<br>';
                    //echo $objWorksheet->getHighestRow() . '<br>';
                    //********************** стилизация листа *********
                    $style_header = array(
                        // Заполнение цветом
                        'fill' => array(
                            'type' => PHPExcel_STYLE_FILL::FILL_SOLID,
                            'color'=>array(
                                'rgb' => 'FFFFFF'
                            )
                        ),
                    );
                    //$last_column = $objWorksheet->getHighestColumn()-1;
                    $objWorksheet->getStyle('A6:S'.$last_row)->applyFromArray($style_header);
                    // задаем автовысоту строки
                    for($i=1; $i<=$last_row; $i++) {
                        if($i !== 5) {
                            $objWorksheet->getRowDimension($i)->setRowHeight(-1);
                        }
                    }
*/


	// Формируем и записываем полный маршрут

        $objWorksheet->mergeCells('A2:I3');

	$highestRow = $objWorksheet->getHighestRow();
            $highestColumn = $objWorksheet->getHighestColumn();
		$highestColumn = 'B';
            $headingsArray = $objWorksheet->rangeToArray('B1:'.$highestColumn.'1',null, true, true, true);
            $headingsArray = $headingsArray[1];
            $r = -1;
            $namedDataArray = array();
            for ($row = 2; $row <= $highestRow; ++$row) {
                $dataRow = $objWorksheet->rangeToArray('B'.$row.':'.$highestColumn.$row,null, true, true, true);
                if ((isset($dataRow[$row]['B'])) && ($dataRow[$row]['B'] > '')) {
                    ++$r;
                    foreach($headingsArray as $columnKey => $columnHeading) {
                        //$namedDataArray[$r][$columnHeading] = $dataRow[$row][$columnKey];
			if ($dataRow[$row][$columnKey] == 'Днепропетровск' OR $dataRow[$row][$columnKey] == 'Дніпро') $dataRow[$row][$columnKey] = 'Днепр';
			if ($dataRow[$row][$columnKey] == 'Жовті Води') $dataRow[$row][$columnKey] = 'Желтые Воды';
			if ($dataRow[$row][$columnKey] == 'пгт.Слобожанское' OR $dataRow[$row][$columnKey] == 'Слобожанское') $dataRow[$row][$columnKey] = 'Подгородное';
			if ($dataRow[$row][$columnKey] == 'Каменское') $dataRow[$row][$columnKey] = 'Днепродзержинск';

			$namedDataArray[] = $dataRow[$row][$columnKey];
                    }
                }
            }

	unset($namedDataArray[0]);
	unset($namedDataArray[1]);

	$result = array_unique($namedDataArray);
	if (is_array($result)) {
		foreach ($result as $route) {
			$allRoute .= ' - ' . trim($route);
		}
		$allRoute = substr($allRoute, 3);    // удаляем первых 3 символа слева
	} else {
		$allRoute = trim($result);
	}

	$objWorksheet->getStyle('A2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $objWorksheet->getStyle('A2')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objWorksheet->getStyle('A2')->getAlignment()->setWrapText(true);

	$objWorksheet->setCellValue('A2', $allRoute);
	unset($allRoute);
	unset($result);
	unset($namedDataArray);

	// Удаляем лишние надписи под таблицей (после Итого:)

                    $i = 0;
                    $flag = true;
                    $rowIterator = $objWorksheet->getRowIterator();
                    foreach ($rowIterator as $row) {
                        if ($flag == false) {
                            break;
                        }
                        // Получили ячейки текущей строки и обойдем их в цикле
                        $i++;
                        $cellIterator = $row->getCellIterator();
                        //print_r($cellIterator);
                        foreach ($cellIterator as $cell) {
                            if (stripos($cell->getCalculatedValue(), 'Итого:') !== FALSE) { //нашли последнюю строку
                                $flag = false;
                                $last_column = $cell->getColumn();
                                $last_row = $cell->getRow();
                                break;
                            }
                        }
                    }
                    $row_id = $i + 1; // deleted row id
                    $number_rows = 50; // number of rows count
                    if ($objWorksheet != NULL) {
                        if ($objWorksheet->removeRow($row_id, $number_rows)) {
                        }
                    }


	// *************************************


                    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
                    $objWriter->save(DIR_SAVE_EDIT_ATTACH . '/' . $filename);
                    $csvWirter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'CSV');
                    $csvWirter->save(DIR_SAVE_CSV . '/' . $filename);
                    $kol++;

         	        $objPHPExcel->__destruct();
                    unset($objWriter);
                    $objWorksheet->__destruct();
                    unset($objWorksheet);
                    unset($objPHPExcel);

                }
            }

            unset($objReader);
            unset($objWriter);
            unset($pExcel);
            unset($objPHPExcel);
            unset($rowIterator);
            unset($cellIterator);
            unset($csvWirter);
            unset($objWorksheet);
  	        gc_collect_cycles();

//echo 'EDITATTACH: memory usage  after unset === <b>' . memory_get_usage() . '</b><br>';

            //************** создаем общую книгу со всеми листами
            //****************************************************
            return $arr_edit_attach; //возвращаем имена файлов и день доставки в массиве
        }
        if($type == 'fsl') {
            $arr_edit_attach = Array();
            //уаляем все файлы из папки edit
            deleteAllFilesFromDirectory(DIR_SAVE_EDIT_ATTACH);
            //-------------------------------
            $kol = 0; //колво файлов
            $all_files = glob(DIR_ATTACH_SAVE .'/*.xlsx');
            $files = delPathFromNameFile($all_files);
            $pExcel = new PHPExcel();
            $objReader = PHPExcel_IOFactory::createReader('Excel2007');
            foreach ($files as $filename) {
                //$xls = PHPExcel_IOFactory::load(DIR_ATTACH_SAVE . '/' . $filename);

            }
            unset($objReader);
            unset($objWriter);
            //************** создаем общую книгу со всеми листами
            //****************************************************
            return $arr_edit_attach; //возвращаем имена файлов и день доставки в массиве

        }
    }


    //************************************************************
    //*******************  создание общей книги ML ***************
    //************************************************************
    public function mlCreate($ml_name) {

        $drivers_mail = Array();
        $drivers = file('config/drivers.txt', FILE_IGNORE_NEW_LINES); // получаем список водителей в массив
//PHPExcel_Settings::setCacheStorageMethod(PHPExcel_CachedObjectStorageFactory::cache_to_sqlite);
        $pExcel = new PHPExcel();
        //переносим лист в книгу
        $objReader = PHPExcel_IOFactory::createReader('Excel5');
        $files = scandir(DIR_SAVE_EDIT_ATTACH);
        unset($files[0], $files[1]); //удаляем знаки '.' и '..' из массива
        $i = 1;
        $arr_list = [];
        $list = 1;
        $arr_titles = array();
        foreach($files as $kk => $fname) {
            if(substr($fname, 0, 2) == 'ml') { //им начинается c ml, значит пишем в книгу
                if(strripos($fname, 'ReestrVozvratov') === FALSE) { //не реестр возвратов
                    // получим имя дл листа (это фамили водителя - ячейка R6C1)
                    $file1 = $objReader->load(DIR_SAVE_EDIT_ATTACH. '/' . $fname);
                    $file1->setActiveSheetIndex(0);
                    $sheet = $file1->getActiveSheet();
                    $title = $this->readName($sheet->getCell('A5')->getValue()); // Задаем имя листа = имени водителя
		            if(empty($title)) { // по какой-то причине не можем прочитать значение в title
			            $title = 'NA';
		            }

		            echo $kk . ' Лист с именем <b>' . $title . '</b> добавлен в книгу ( ' . $fname . ' )<br>';
                    //добавляем запланированного водителя в массив
                    $drivers_mail[$list]['driver'] = $title;
                    $drivers_mail[$list]['ml'] = '1'; //значит водитель с маршрутом
                    //надо удалить из массива имен $drivers значение = $title
                    if(($key = array_search($title,$drivers)) !== FALSE){
                        unset($drivers[$key]);
                    }

                    if (in_array($title, $arr_titles)) {
                        //такой водитель уже есть, надо переименовать
                        $title .= '51';
			        echo 'new title = ' . $title . '<br>';
                    }

                    $arr_titles[] = $title;
                    if(in_array($title, $arr_list)) {
                        // маршрут на такого водителя уже занесен в базу,
                        // т.е. получаются маршруты на разные дни
                        // не заносим, создаем массив с незанесенными маршрутами
                        // потом занесем их в отдельный файл
                    } else {
                        $arr_list[] = $title; //бывает, что на одного водителя больше одного листа
                        $sheet = $file1->getActiveSheet()->setTitle($title);
                        $pExcel->addExternalSheet($sheet);
                        $pExcel->setActiveSheetIndex($list);
                        $new_sheet = $pExcel->getActiveSheet();
                        // задаем ширину столбцов
                        $new_sheet->getColumnDimension('A')->setWidth(20);
                        $new_sheet->getColumnDimension('B')->setWidth(11);
                        $new_sheet->getColumnDimension('C')->setWidth(25);
                        $new_sheet->getColumnDimension('D')->setWidth(15);
                        $new_sheet->getColumnDimension('E')->setWidth(15);
                        $new_sheet->getColumnDimension('F')->setWidth(3);
                        $new_sheet->getColumnDimension('G')->setWidth(3);
                        $new_sheet->getColumnDimension('H')->setWidth(17);
                        $new_sheet->getColumnDimension('I')->setWidth(7);
                        $new_sheet->getColumnDimension('J')->setWidth(5);
                        $new_sheet->getColumnDimension('K')->setWidth(5);
                        $new_sheet->getColumnDimension('L')->setWidth(5);
                        $new_sheet->getColumnDimension('M')->setWidth(7);

                        /*
                        $new_sheet->getColumnDimension('N')->setWidth(7);
                        $new_sheet->getColumnDimension('O')->setWidth(5);
                        $new_sheet->getColumnDimension('P')->setWidth(6);
                        $new_sheet->getColumnDimension('Q')->setWidth(6);
                        $new_sheet->getColumnDimension('R')->setWidth(6);
                        $new_sheet->getColumnDimension('S')->setWidth(7);
                        */
                        for($a=1; $a<=100; $a++) {
                            if($a !== 5) {
                                $new_sheet->getRowDimension($a)->setRowHeight(-1);
                            }
                        }
			// Высота строки шапки таблицы
		        $new_sheet->getRowDimension(6)->setRowHeight(33);
			// ----------------------------
                        $i++;
                        $list++;
			unset($new_sheet);
                    }
			unset($file1);
			unset($sheet);
			//unset($title);
                }
            }
        }
        echo '<hr>';
        foreach ($drivers as $driver) {
            $pExcel->createSheet($list);
            $pExcel->setActiveSheetIndex($list);
            $pExcel->getActiveSheet()->setTitle($driver);
            $list++;
            echo "<b style='color:#ff081b'>Добавлен лист без маршрута для водителя " . $driver . "</b><br>";
            //добавляем запланированного водителя в массив
            $drivers_mail[$list]['driver'] = $driver;
            $drivers_mail[$list]['ml'] = '0'; //значит водителm без маршрута

        }

        echo '<hr>';
        $pExcel->removeSheetByIndex(0);
        //сохраняем файл
        $objWriter = PHPExcel_IOFactory::createWriter($pExcel, 'Excel5');
        $objWriter->save(DIR_SAVE_ML . '/' .$ml_name);

	    $pExcel->__destruct();
        $objReader = null;
        $objWriter = null;
        $file1 = null;
        $new_sheet = null;
        $sheet = null;
        $pExcel = null;

        unset($objReader);
        unset($objWriter);
        unset($file1);
        unset($new_sheet);
        unset($sheet);
        unset($pExcel);
	    gc_collect_cycles();
        if(file_exists(DIR_SAVE_ML . '/' . $ml_name)) {
            return $drivers_mail;
        } else {
            return 0;
        }
    }


    public function editCart($files, $type)
    {
        $edit_files = [];
        if ($type == 'cart') {
            deleteAllFilesFromDirectory(DIR_EDIT_SAVE_CART);
            //$all_files = glob(DIR_ATTACH_SAVE_CART . '/*.xls');
            //$files = delPathFromNameFile($all_files);
            $pExcel = new PHPExcel();
            $objReader = PHPExcel_IOFactory::createReader('Excel5');
            foreach ($files as $filename) {
                $objPHPExcel = $objReader->load(DIR_ATTACH_SAVE_CART . '/' . $filename);
                $objWorksheet = $objPHPExcel->getActiveSheet();
                $rowIterator = $objWorksheet->getRowIterator();

                $box = [];
                $coordinates = '';
                $strBox = '';

                foreach ($rowIterator as $row) {
                    // Получили ячейки текущей строки и обойдем их в цикле
                    $cellIterator = $row->getCellIterator();
                    foreach ($cellIterator as $cell) {
                        if (trim($cell->getValue()) == 'Накладные') {
                            if (!empty($box[$coordinates]) && !empty($coordinates)) {
                                $strBox = '';
                                foreach ($box[$coordinates] as $key => $val) {
                                    if ($key == 'box') {
                                        $strBox .= 'BOX - ' . $val . PHP_EOL;
                                    } elseif ($key == 'box_big') {
                                        $strBox .= 'BOX (большой) - ' . $val . PHP_EOL;
                                    } elseif ($key == 'box_small') {
                                        $strBox .= 'BOX (маленький) - ' . $val . PHP_EOL;
                                    }

                                }
                                $objWorksheet->setCellValue($coordinates, (string)$strBox);
                                $arHeadStyle = array(
                                    'font'  => array(
                                        //'bold'  => true,
                                        //'color' => array('rgb' => '778899'),
                                        'size'  => 11,
                                        //'name'  => 'Verdana'
                                    ));

                                # применение стилей к ячейкам
                                $objWorksheet->getStyle($coordinates)->applyFromArray($arHeadStyle);
                            }
                            // получаем данные по ячейкам
                            //$beginRow = $row->getRowIndex(); //начало накладной клиента
                            //$beginCell = $cell->getColumn(); // получаем столбец (например 'D')
                            //$coordDataRow = $beginRow + 3;
                            //$coordDataCell = 10;
                            $coordinates = 'J' . (string)($row->getRowIndex() + 3);
                            $objWorksheet->setCellValue($coordinates, '');
                            //echo "Координаты - $coordinates<br>";
                            //$objWorksheet->setCellValue($coordinates, 'УРЯЯЯ!!!');

                        } else {
                            if (trim($cell->getValue()) == 'BOX') {
                                $box[$coordinates]['box'] = $box[$coordinates]['box'] + 1;
                            } elseif (trim($cell->getValue()) == 'BOX (Маленький)') {
                                $box[$coordinates]['box_small'] = $box[$coordinates]['box_small'] + 1;
                            } elseif (trim($cell->getValue()) == 'BOX(большой)') {
                                $box[$coordinates]['box_big'] = $box[$coordinates]['box_big'] + 1;
                            }

                        }
                        //$row->getRowIndex(); // получаем номер строки
                        //$cell->getColumv(); // получаем столбец (например 'D')
                    }
                }

                //var_dump($box);

                $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
                $objWriter->save(DIR_EDIT_SAVE_CART . '/' . $filename);
                //break;
                $edit_files[] = DIR_EDIT_SAVE_CART . '/' . $filename;
                unset($objWriter);
                unset($rowIterator);
                unset($objWorksheet);
                unset($objPHPExcel);
                gc_collect_cycles();
            }
        }
        unset($objReader);
        unset($pExcel);
        gc_collect_cycles();
        return $edit_files;
    }

}
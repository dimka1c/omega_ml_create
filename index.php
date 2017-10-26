<?php
    session_start();
//echo 'memory usage start script === ' . round(memory_get_usage()/1024/1024, 2) . 'Mb<br>';
    date_default_timezone_set('Europe/Kiev');

    require_once 'config/config.php';
    require_once 'class/functions.php';

/*    require_once 'class/emailClass.php';
    require_once 'class/workExcel.php';
    require_once 'class/PHPExcel.php';
*/

    // делаем автозагрузку классов
    if(!function_exists('classOmegaAutoLoader')){
        function classOmegaAutoLoader($class){
            $classFile=$_SERVER['DOCUMENT_ROOT'].'/class/'.$class.'.php';

            if(is_file($classFile)&&!class_exists($class)) {
		include $classFile;

		}
        }
    }
    spl_autoload_register('classOmegaAutoLoader');
/*
	echo 'Ошибка сервера<br>';
	echo 'Формирование ML не выполнено <br>';
	exit;
*/

        $file = 'log/log.txt';
        // Открываем файл для получения существующего содержимого
        $current = file_get_contents($file);
        // Добавляем нового человека в файл
        //$current = file_get_contents($file);
        // Добавляем нового человека в файл
        $current = date('d-m-Y H:i:s') . PHP_EOL;
        $current .= 'refer = ' . $_SERVER['HTTP_REFERER'] . PHP_EOL;
        $current .= 'browser = ' . $_SERVER['HTTP_USER_AGENT'] . PHP_EOL;
        $current .= 'ip = ' . $_SERVER['REMOTE_ADDR'] . PHP_EOL;
        $current .= 'host = ' . $_SERVER['REMOTE_HOST'] . PHP_EOL;
        $current .= 'uri = ' . $_SERVER['REQUEST_URI'] . PHP_EOL;
        $current .= '--------------------------------------------------------------------'. PHP_EOL;
	// Пишем содержимое обратно в файл
        file_put_contents($file, $current, FILE_APPEND);
        unset($current);

	//include 'error/404.htm'; 
	//exit;



    //$host = MAIL_HOST_SSQ_PP_UA;
    $host = MAIL_HOST_GOOGLE;
    //$host = MAIL_HOST_GOOGLE;
    // получаем письма из почтового ящика
    // в идеале должно быть одно письмо с определенным subject
    // однако может быть несколько писем
    // надо выбрать  нужное письмо, остальные  письма удалить
    $mail = new emailClass;

    // получаем все письма из ящике

    $mail_data = $mail->receiveEmail($host);

//debug($mail_data);

    // ищем нужное нам письмо,
    // критерий поиска :
    // - только определенный отправитель
    // - тема письма должна соответствовать шаблону "ml/16_11_2016/create"
    // функция вернет массив с письмом или письмами, если несколько
    // или вернет пустой массив, если нет нужных писем
    // array (size=3)
    //      'uid' => int 7
    //      'from' => string 'dima@udt.dp.ua' (length=14)
    //      'subj' => string 'ml/16_11_2016/create' (length=20)
    if(empty($mail_data)) { //если нет писем в папке, то просто завершием работу скрипта
        echo 'нет писем в почтовом ящике <br>';
        unset($mail_data);
        unset($mail);
        exit;
    }
    $yes_mail = $mail->findRightMail($mail_data);
    if (empty($yes_mail)) {
	echo 'Письма для формирования файла не обнаружены. Возможно не верно указана тема письма.<br>';
	echo 'В ящике находятся письма со следующими темами:<br>';
	foreach ($mail_data as $arr) {
	    echo $arr['subj'] . '<br>';
	}
    exit;
    }


    if(!empty($yes_mail)) {
        //если есть письма
        // проверем вложения и сохраняем на диск

        foreach ($yes_mail as $arr) {

            if ($arr['type'] == 'cart') {
                deleteAllFilesFromDirectory(DIR_ATTACH_SAVE_CART);
                $attach = $mail->loadAttach($host, $arr['uid'], DIR_ATTACH_SAVE_CART, $arr['type']);
                if(chekFilesInFolder($attach, DIR_ATTACH_SAVE_CART)) {
                    $excel = new workExcel;
                    $edit_attach = $excel->editCart($attach, $arr['type']);
                    if (!empty($edit_attach)) {
                        // вернулся массив, файлы сохранены
                        // отправляем готовый файл отправителю
                        foreach ($edit_attach as $file) {
                            echo '<pre>' . print_r($file) . '</pre>';
                            $send_mail_cart = $mail->sendMailCart($arr['from'], $arr['subj'], $file);
                            if($send_mail_cart) { // письмо отправлено, удаляем письмо из почтового ящика
                                echo 'Письмо от ' . $arr['from']. ' с темой ' . $arr['subj'] . ' сформировано и отправлено на адрес: ' . $arr['from'] . '<br>';
                                $del_mail = $mail->delMail($host, $arr['uid']);
                                echo 'Письмо удалено с сервера<br>';
                            }
                        }
                    }
                }

            } elseif($arr['type'] == 'ml') {

                deleteAllFilesFromDirectory(DIR_ATTACH_SAVE);
                $attach = $mail->loadAttach($host, $arr['uid'], DIR_ATTACH_SAVE, $arr['type']);
		if (empty($attach)) {
		    echo 'Ошибка! Файлы не были получены с почтового ящика.<br>';
		    echo 'Однако на почте файлы присутствуют:<br>';
		    foreach($yes_mail as $mail_files) {
			echo '<h4> - письмо №' . $mail_files['uid'] . '. Отправитель: ' . $mail_files['from'] . '. Тема письма: ' . $mail_files['subj'] . '</h4><br>';
		    }
		    echo 'Работа программы завершена. <br>';
		    exit;
		}

                // сверяем массив полученных файлов с сохраненными файлами в папке
                if(chekFilesInFolder($attach, DIR_ATTACH_SAVE)) {
                    //файлы все есть, начинаем создание общей книги
                    // сначала редактируем файлы по очереди
                    $excel = new workExcel;
                    // редактируем файлы xls (удаляем строки, изменяем ширину строк и столбцов)
                    // и сохраняем на диск
                    $edit_attach = $excel->editAttachFiles($attach, $arr['type']);
                    // создаем общую книгу ml
                    $ml_file_name = getMlFileName($arr['subj']); // из subj получаем имя файла ml/07_12_2016/create => ml_07_12_2016.xls
                    $ml_create = $excel->mlCreate($ml_file_name); // вернет массив с водителями (запланированные и нет), 0 - файл не создан
                    // $ml_create - массив с запланированными и незапланированными водителями
                    // если $ml_create = 0 - ошибка создания файла ml
                    if((!empty($ml_create)) && ($ml_create !== 0)) { // вернулся массив
                        // отправляем готовый файл отправителю
			$send_mail = $mail->sendMail($arr['from'], $arr['subj'], $ml_file_name, $attach, $ml_create);
			//$send_mail = true;
                        if($send_mail) { // письмо отправлено, удаляем письмо из почтового ящика
                            echo 'Письмо от ' . $arr['from']. ' с темой ' . $arr['subj'] . ' сформировано и отправлено на адрес: ' . $arr['from'] . '<br>';
			$del_mail = $mail->delMail($host, $arr['uid']);
                            echo 'Письмо удалено с сервера<br>';
                        }
                        // заносим данные из отредактированных книг в БД
                    }
                }
            }
            unset($edit_attach);
            unset($ml_create);
            unset($excel);
	        gc_collect_cycles();
            echo 'memory usage  end script === ' . round(memory_get_usage()/1024/1024, 2) . 'Mb<br>';
            echo '<hr>';

        }

    } else {
        // писем, соответствующих шгаблону нет
        exit;
    }

?>

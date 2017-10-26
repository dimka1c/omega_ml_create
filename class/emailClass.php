<?php

class emailClass
{

    private $part_array = Array();
	private $structure;
	private $inbox;
	private $attach;
	private $file_attachment;
	private $structureKey = -1;


    private function create_part_array($structure, $prefix="") {
        //print_r($structure);
        global $part_array;
        if (sizeof($structure->parts) > 0) {    // There some sub parts
            foreach ($structure->parts as $count => $part) {
                $this->add_part_to_array($part, $prefix.($count+1), $part_array);
            }
        }else{    // Email does not have a seperate mime attachment for text
            $part_array[] = array('part_number' => $prefix.'1', 'part_object' => $obj);
        }
        return $part_array;
    }

    // Sub function for create_part_array(). Only called by create_part_array() and itself.
    private function add_part_to_array($obj, $partno, $part_array) {
        global $part_array;
        $part_array[] = array('part_number' => $partno, 'part_object' => $obj);
        if ($obj->type == 2) { // Check to see if the part is an attached email message, as in the RFC-822 type
            //print_r($obj);
            if (sizeof($obj->parts) > 0) {    // Check to see if the email has parts
                foreach ($obj->parts as $count => $part) {
                    // Iterate here again to compensate for the broken way that imap_fetchbody() handles attachments
                    if (sizeof($part->parts) > 0) {
                        foreach ($part->parts as $count2 => $part2) {
                            $this->add_part_to_array($part2, $partno.".".($count2+1), $part_array);
                        }
                    }else{    // Attached email does not have a seperate mime attachment for text
                        $part_array[] = array('part_number' => $partno.'.'.($count+1), 'part_object' => $obj);
                    }
                }
            }else{    // Not sure if this is possible
                $part_array[] = array('part_number' => $prefix.'.1', 'part_object' => $obj);
            }
        }else{    // If there are more sub-parts, expand them out.
            if (sizeof($obj->parts) > 0) {
                foreach ($obj->parts as $count => $p) {
                    $this->add_part_to_array($p, $partno.".".($count+1), $part_array);
                }
            }
        }
    }


    // получаем все письма с почтового адреса $host (gmail или yandex)
    // возвращаем uid-письма, адресат(от кого пришло), заголовок (subj)
    public function receiveEmail($host)
    {
        if($host == MAIL_HOST_YANDEX) {
            $username = MAIL_USER_NAME_YANDEX;
            $password = MAIL_USER_PASSWORD_YANDEX;
        } elseif ($host == MAIL_HOST_GOOGLE) {
            $username = MAIL_USER_NAME_GOOGLE;
            $password = MAIL_USER_PASSWORD_GOOGLE;
        } elseif ($host == MAIL_HOST_SSQ_PP_UA) {
            $username = MAIL_USER_NAME_SSQ_PP_UA;
            $password = MAIL_USER_PASSWORD_SSQ_PP_UA;
        } elseif ($host == MAIL_HOST_YAHOO) {
            $username = MAIL_USER_NAME_YAHOO;
            $password = MAIL_USER_PASSWORD_YAHOO;
        }

        set_time_limit(60000);
        $ml = imap_open($host, $username, $password) or die('Cannot connect to: ' . $host . '. <b>ERROR</b>' . imap_last_error());
        if($ml) {
            $n = imap_num_msg($ml); //колво писем в ящике
            if ($n > 0) {
                $uids = array();
                $mails = array();
                for ($i=1; $i<=$n; $i++) {
                    $h = imap_header($ml, $i);
                    $h = $h->from;
                    foreach ($h as $k => $v) {
                        $mailbox = $v->mailbox;
                        $host = $v->host;
                        $personal = $v->personal;
                        $email = $mailbox . '@' . $host;
                        //return $mailbox . " / " . $host . " / " . $personal . " / " . $email;
                    }
                    $headerArr = imap_headerinfo($ml, $i);

                    $uid = imap_uid($ml, $headerArr->Msgno);

                    $mails[] = $email;
                    $uids[] = $uid;
                    $arr[$i]['uid'] = $uid;
                    $arr[$i]['email'] = $email;

                    $s = imap_fetch_overview($ml, $uid, FT_UID);
                    foreach ($s as $k => $v) {
                        $subj = imap_utf8($v->subject);
                        //$subj = $v->subject;
                        $arr[$i]['subj'] = $subj;
                    }
                }
            }
        }
        imap_close($ml);
        unset($ml);
        unset($headerArr);
//var_dump($arr);
        return $arr;

    }


    // функция разбора поля subject
    private static function explodeSubject($subj) {
        $subj = trim($subj);
        if (preg_match(PATTERN_SUBJ, $subj)) {
            // это обработка ml
            return 1;
        }
        if (preg_match(PATTERN_CART, $subj)) {
            // это обработка карты загрузки
            return 1;
        }
        return 0;
    }

    // возвращвет тип обрабатываемого файла (ml, fsl) из subj
    private static function getType($subj) {
        $type = explode('/', $subj);
        if(isset($type[0])) {
            return $type[0];
        }
        return 'NA'; //тип сообщения неопределен
    }

    /************************************************************************
     * ищем нужное письмо среди массива писем или одного письма (если в ящике одно письмо)
     * и возвращаем uid-письма, email, и дату ml
     */
    public function findRightMail($mail_data)
    {

        $mails = array();
        $i = 0; //счетчик правильных писем
        foreach ($mail_data as $arr) {
            $uid = $arr['uid'];
            $from = $arr['email'];
            $subj = $arr['subj'];
            //сначала ищем по отправителям
            if (($from == 'dima@udt.dp.ua') ||
                ($from == 'omegadp2017@yandex.ua') ||
                ($from == 'omega_dnepr@yahoo.com') ||
                ($from == 'omega@ssq.pp.ua') ||
                ($from == 'omega@udt.dp.ua') ||
                ($from == 'eremtsov.maksim@omega-auto.biz') ||
                ($from == 'turchin.vladimir@omega-auto.biz')) {
                    // адресат совпал, ищем по subject
                    if(self::explodeSubject($subj)) {
                        //subj правильный
                        $mails[$i]['uid'] = $uid;
                        $mails[$i]['from'] = $from;
                        $mails[$i]['subj'] = $subj;
                        $mails[$i]['type'] = self::getType($subj);
                        $i++;
                    }
            }
        }
        return $mails;
    }

    //*****************************************
    private static function isFileName($fname) {

        $f = substr($fname, 0, 2);
        if($f == 'ml') {
            return 1;
        }
        return 0;
    }

    /************************************************************************
     * сохраняем вложения на диск в папку $path
     * и возвращаем uid-письма, email, и дату ml
     */
    public function loadAttach($host, $uid, $path, $type)
    {

        $arr_attach = Array();

        if($host == MAIL_HOST_YANDEX) {
            $username = MAIL_USER_NAME_YANDEX;
            $password = MAIL_USER_PASSWORD_YANDEX;
        } elseif ($host == MAIL_HOST_GOOGLE) {
            $username = MAIL_USER_NAME_GOOGLE;
            $password = MAIL_USER_PASSWORD_GOOGLE;
        } elseif ($host == MAIL_HOST_SSQ_PP_UA) {
            $username = MAIL_USER_NAME_SSQ_PP_UA;
            $password = MAIL_USER_PASSWORD_SSQ_PP_UA;
        } elseif ($host == MAIL_HOST_YAHOO) {
            $username = MAIL_USER_NAME_YAHOO;
            $password = MAIL_USER_PASSWORD_YAHOO;
        }
        set_time_limit(40000);

        $inbox = imap_open($host, $username, $password) or die('Cannot connect to mailbox: ' . imap_last_error());
        $structure = imap_fetchstructure($inbox, $uid, FT_UID);
        $part_array = $this->create_part_array($structure);
        //$header = imap_fetchbody($inbox, $uid, '0', FT_UID);

        foreach ($part_array as $key => $attach) {
//debug($attach['part_object']);
            if (($attach['part_object']->type == 3) &&
                (strtoupper($attach['part_object']->disposition) == 'ATTACHMENT') &&
                ($attach['part_object']->ifdparameters == 1) &&
                ((strtoupper($attach['part_object']->dparameters[0]->attribute) == 'FILENAME') OR 
		 (strtoupper($attach['part_object']->dparameters[1]->attribute) == 'FILENAME'))
            ) {
		if (isset($attach['part_object']->dparameters)) {
		    foreach ($attach['part_object']->dparameters as $key => $fname) {
			if (strtoupper($fname->attribute) == 'FILENAME') {
			    $this->structureKey = $key;
			    continue;
			}
		    }
		}

//echo 'нашли вложение ' . $this->structureKey . '<br>';

		if ($this->structureKey == -1) { 
		    echo 'Не возможно найти вложения. Программа завершена <br>';
		    exit;
		}
                //нашли тело, которое надо вытащить
                //echo 'number = ' . $attach['part_number'] . ' / file = ' . $attach['part_object']->dparameters[0]->value . '<br>';
                //$filename = $attach['part_object']->dparameters[0]->value;
		$filename = $attach['part_object']->dparameters[$this->structureKey]->value;
                if ($type == 'fsl') {
                    $filename = str2url(iconv_mime_decode($filename));
                    $file_attachment = imap_fetchbody($inbox, $uid, $attach['part_number'], FT_UID);
                    if ($attach['part_object']->encoding == 3) {
                        $file_attachment = base64_decode($file_attachment);
                    } elseif ($attach['part_object']->encoding == 4) {
                        $file_attachment = quoted_printable_decode($file_attachment);
                    }
                    $fp = fopen(DIR_ATTACH_SAVE . '/' . $filename, "w+");
                    fwrite($fp, $file_attachment);
                    fclose($fp);
                    unset($fp);
                    $arr_attach[] = $filename;
                    //echo 'Save file  ' . $filename . '<br>';
                }
                if ($type === 'cart') {
                    $filename = str2url(str_replace(" ", "_", iconv_mime_decode($filename)));
                    $file_attachment = imap_fetchbody($inbox, $uid, $attach['part_number'], FT_UID);
                    if ($attach['part_object']->encoding == 3) {
                        $file_attachment = base64_decode($file_attachment);
                    } elseif ($attach['part_object']->encoding == 4) {
                        $file_attachment = quoted_printable_decode($file_attachment);
                    }
                    $fp = fopen(DIR_ATTACH_SAVE_CART . '/' . $filename, "w+");
                    fwrite($fp, $file_attachment);
                    fclose($fp);
                    unset($fp);
                    $file_attachment = null;
                    unset($file_attachment);
                    $arr_attach[] = $filename;
                    //echo 'Save file  ' . $filename . '<br>';
                }
                if (self::isFileName($filename) == 1) {
                    $file_attachment = imap_fetchbody($inbox, $uid, $attach['part_number'], FT_UID);
                    if ($attach['part_object']->encoding == 3) {
                        $file_attachment = base64_decode($file_attachment);
                    } elseif ($attach['part_object']->encoding == 4) {
                        $file_attachment = quoted_printable_decode($file_attachment);
                    }
                    $fp = fopen(DIR_ATTACH_SAVE . '/' . $filename, "w+");
                    fwrite($fp, $file_attachment);
                    fclose($fp);
                    unset($fp);
                    $file_attachment = null;
                    unset($file_attachment);
                    $arr_attach[] = $filename;
                    //echo 'Save file  ' . $filename . '<br>';
                }
            } //endif
        }
        imap_close($inbox);

        $file_attachment = null;
        unset($file_attachment);
	    $filename = null;
	    unset($filename);
        $attach = null;
        unset($attach);
        $part_array = null;
        unset($GLOBALS['part_array']);
        $structure = null;
        unset($structure);
        $inbox = null;
        unset($inbox);

        return $arr_attach;
    }


    // отправка почтового сообщения
    public function sendMailAttachment($mailTo, $from, $subject, $message, $file = false)
    {
        $separator = "---"; // разделитель в письме
        // Заголовки для письма
        $headers = "MIME-Version: 1.0\r\n";
        $headers .= "From: $from\nReply-To: $from\n"; // задаем от кого письмо
        $headers .= "Content-Type: multipart/mixed; boundary=\"$separator\""; // в заголовке указываем разделитель
        // если письмо с вложением
        if ($file) {
            $bodyMail = "--$separator\n"; // начало тела письма, выводим разделитель
            $bodyMail .= "Content-type: text/html; charset='utf-8'\n"; // кодировка письма
            $bodyMail .= "Content-Transfer-Encoding: quoted-printable"; // задаем конвертацию письма
            $bodyMail .= "Content-Disposition: attachment; filename==?utf-8?B?" . base64_encode(basename($file)) . "?=\n\n"; // задаем название файла
            $bodyMail .= $message . "\n"; // добавляем текст письма
            $bodyMail .= "--$separator\n";
            $fileRead = fopen($file, "r"); // открываем файл
            $contentFile = fread($fileRead, filesize($file)); // считываем его до конца
            fclose($fileRead); // закрываем файл
            $bodyMail .= "Content-Type: application/octet-stream; name==?utf-8?B?" . base64_encode(basename($file)) . "?=\n";
            $bodyMail .= "Content-Transfer-Encoding: base64\n"; // кодировка файла
            $bodyMail .= "Content-Disposition: attachment; filename==?utf-8?B?" . base64_encode(basename($file)) . "?=\n\n";
            $bodyMail .= chunk_split(base64_encode($contentFile)) . "\n"; // кодируем и прикрепляем файл
            $bodyMail .= "--" . $separator . "--\n";
            // письмо без вложения
        } else {
            $bodyMail = $message;
        }
        $result = mail($mailTo, $subject, $bodyMail, $headers); // отправка письма
        unset($bodyMail);
        return $result;

    }

    public function sendMail($to, $subj, $file, $attach, $drivers)
    {

        $file = DIR_SAVE_ML . '/' . $file; // файл аттач
        $mailTo = $to; //"dima@udt.dp.ua"; // кому
        //$mailTo = "dima@udt.dp.ua"; // кому
        $from = MAIL_USER_NAME_SSQ_PP_UA; //"omega@udt.dp.ua"; // от кого
        $subject = 'Сформированный файл по запросу ' . $subj; //"Test file"; // тема письма

        if(isset($drivers)) {
            $message = '<br>Информация по водителям:</b><br><br>';
            foreach ($drivers as $driver) {
                if($driver['ml'] == 1) {
                    $message .= $driver['driver'] . ' запланирован в маршруте<br>';
                } elseif ($driver['ml'] == 0) {
                    $message .= "<b style='color:#ff081b'>" . $driver['driver'] . " не запланирован в маршруте</b><br>";
                }
            }
        }


        $message .= "<br>Обработанные файлы:<br>";
        foreach ($attach as $key => $val) {
            $message .= $key + 1 .' - ' . $val . "<br>";
        }
        //print_r($message);
        //$message = //"Тестовое письмо с вложением"; // текст письма
        $r = $this->sendMailAttachment($mailTo, $from, $subject, $message, $file); // отправка письма c вложением
        //echo ($r) ? 'Письмо отправлено' : 'Ошибка. Письмо не отправлено!'
        return $r;
    }

    // удаляем письмо из почтового ящика
    public function delMail($host, $uid) {
        $arr_attach = Array();
        if($host == MAIL_HOST_YANDEX) {
            $username = MAIL_USER_NAME_YANDEX;
            $password = MAIL_USER_PASSWORD_YANDEX;
        } elseif ($host == MAIL_HOST_GOOGLE) {
            $username = MAIL_USER_NAME_GOOGLE;
            $password = MAIL_USER_PASSWORD_GOOGLE;
        } elseif ($host == MAIL_HOST_SSQ_PP_UA) {
            $username = MAIL_USER_NAME_SSQ_PP_UA;
            $password = MAIL_USER_PASSWORD_SSQ_PP_UA;
        } elseif ($host == MAIL_HOST_YAHOO) {
            $username = MAIL_USER_NAME_YAHOO;
            $password = MAIL_USER_PASSWORD_YAHOO;
        }

        $inbox = imap_open($host, $username, $password) or die('Cannot connect to mailbox: ' . imap_last_error());
//	imap_mail_move ($inbox, $uid, 'Trash'); 
	
        imap_delete ($inbox, $uid, FT_UID);
        imap_expunge ($inbox);

        imap_close($inbox);

    }


    public function sendMailCart($from, $subj, $file)
    {
        //$mailTo = 'dima@udt.dp.ua';
        $mailTo = $from;
        $separator = "---"; // разделитель в письме
        // Заголовки для письма
        $headers = "MIME-Version: 1.0\r\n";
        $headers .= "From: $from\nReply-To: $from\n"; // задаем от кого письмо
        $headers .= "Content-Type: multipart/mixed; boundary=\"$separator\""; // в заголовке указываем разделитель
        // если письмо с вложением
        if ($file) {
            $bodyMail = "--$separator\n"; // начало тела письма, выводим разделитель
            $bodyMail .= "Content-type: text/html; charset='utf-8'\n"; // кодировка письма
            $bodyMail .= "Content-Transfer-Encoding: quoted-printable"; // задаем конвертацию письма
            //echo '<pre>' . print_r(base64_encode(basename($file))) . '</pre>';
            $bodyMail .= "Content-Disposition: attachment; filename==?utf-8?B?" . base64_encode(basename($file)) . "?=\n\n"; // задаем название файла
            $bodyMail .= 'тестовое письмо' . "\n"; // добавляем текст письма
            $bodyMail .= "--$separator\n";
            $fileRead = fopen($file, "r"); // открываем файл
            $contentFile = fread($fileRead, filesize($file)); // считываем его до конца
            fclose($fileRead); // закрываем файл
            $bodyMail .= "Content-Type: application/octet-stream; name==?utf-8?B?" . base64_encode(basename($file)) . "?=\n";
            $bodyMail .= "Content-Transfer-Encoding: base64\n"; // кодировка файла
            $bodyMail .= "Content-Disposition: attachment; filename==?utf-8?B?" . base64_encode(basename($file)) . "?=\n\n";
            $bodyMail .= chunk_split(base64_encode($contentFile)) . "\n"; // кодируем и прикрепляем файл
            $bodyMail .= "--" . $separator . "--\n";
            // письмо без вложения
        } else {
            $bodyMail = 'тестовое сообщение';
        }
        $result = mail($mailTo, $subj, $bodyMail, $headers); // отправка письма
        unset($bodyMail);
        return $result;
    }

}

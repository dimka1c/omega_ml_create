<?php

class logClass
{

    public static function Client()
    {
        date_default_timezone_set('Europe/Kiev');

echo __DIR__ .'<br>';

        $file = 'log/log.txt';
        // ��������� ���� ��� ��������� ������������� �����������
        $current = file_get_contents($file);
        // ��������� ������ �������� � ����
        $current = file_get_contents($file);
        // ��������� ������ �������� � ����
        $current .= date('Y-m-d H:i:s') . PHP_EOL;
        $current .= 'refer = ' . $_SERVER['HTTP_REFERER'] . PHP_EOL;
        $current .= 'browser = ' . $_SERVER['HTTP_USER_AGENT'] . PHP_EOL;
        $current .= 'ip = ' . $_SERVER['REMOTE_ADDR'] . PHP_EOL;
        $current .= 'host = ' . $_SERVER['REMOTE_HOST'] . PHP_EOL;
        $current .= 'uri = ' . $_SERVER['REQUEST_URI'] . PHP_EOL;
        $current .= '--------------------------------------------------------------------'. PHP_EOL;
	// ����� ���������� ������� � ����
        file_put_contents($file, $current, FILE_APPEND);

    }

}
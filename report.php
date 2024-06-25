<?php
header('Content-type: text/html; charset=utf8');
require_once "src/config.php";
require_once "src/mysql.php";
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\{Alignment};
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use src\config;
use src\mysql;


$config= new config();
$mysql= new mysql($config->mysql);
$reader = new \PhpOffice\PhpSpreadsheet\Reader\Html();
function nod($a, $b) {
    return $a? nod($b%$a, $a) : $b;
}
function nok($a, $b) {
    return $a / nod($a, $b) * $b;
}

$res= $mysql->table("2report")->order(["uID","appPriority","appID"])->get()->getResults();

$spreadsheet = new Spreadsheet();
$styles=[
    'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'wrapText' => true,
    ]
];
$sheet1 = $spreadsheet->getActiveSheet();
$rs=3;
$lettersHead=[
    "A"=>"appID",
    "B"=>"Дата ввода",
    "C"=>"Фамилия",
    "D"=>"Имя",
    "E"=>"Отчество",
    "F"=>"Дата рождения",
    "G"=>"СНИЛС",
    "H"=>"Пол",
    "I"=>"Льготы",
    "J"=>"Email",
    "K"=>"Гражданство",
    "L"=>"Вид документа",
    "M"=>"Серия документа",
    "N"=>"№ документа",
    "O"=>"Кем выдан",
    "P"=>"Дата выдачи",
    "Q"=>"Место рождения",
    "R"=>"Статус",
    "S"=>"Оригинал документов",
    "T"=>"Вид оригинала",
    "U"=>"Оператор ПК",
    "V"=>"Адрес по прописке",
    "W"=>"Адрес текущего проживания",
    "X"=>"Общежитие",
    "Y"=>"Мобильный",
    "Z"=>"Телефон",
    "AA"=>"Тип населенного пунтка",
    "AB"=>"Офиц. Название учебного заведения",
    "AC"=>"Тип учебного заведения",
    "AD"=>"Страна УЗ",
    "AE"=>"Регион УЗ",
    "AF"=>"Изучаемый язык",
    "AG"=>"Год окончания",
    "AH"=>"Вид документа",
    "AI"=>"Серия",
    "AJ"=>"Номер",
    "AK"=>"Дата выдачи",
    "AL"=>"Средний балл атестата",
    "AM"=>"Родство",
    "AN"=>"ФИО",
    "AO"=>"Телефон",
    "AP"=>"Документ",
    "AQ"=>"Документы на льготы",
    "AR"=>"Название испытания",
    "AS"=>"Дисциплина",
    "AT"=>"Балл",
    "AU"=>"Дата испытания",
    "AV"=>"Условие",
    "AW"=>"Конкурсная группа",
    "AX"=>"Название достижения ",
    "AY"=>"Балл достижения",
    "AZ"=>"Приоритет",
    "BA"=>"Статус",
    "BB"=>"Уровень",
    "BC"=>"Форма обучения",
    "BD"=>"Курс",
    "BE"=>"Направление Код",
    "BF"=>"Направление Название",
    "BG"=>"Основание",
    "BH"=>"Номер Л.Д.",
    "BI"=>"Способ подачи",
    "BJ"=>"Подано согласие",
    "BK"=>"№ приказа",
    "BL"=>"Дата приказа",
];
$lettersMax=[
    "A"=>"appID",
    "B"=>"appDate",
    "C"=>"surname",
    "D"=>"name",
    "E"=>"patronymic",
    "F"=>"birthday",
    "G"=>"snils",
    "H"=>"sex",
    "I"=>"benefits",
    "J"=>"email",
    "K"=>"citizenship",
    "L"=>"docType",
    "M"=>"docSerial",
    "N"=>"docNumber",
    "O"=>"docWhoIssued",
    "P"=>"docDateIssued",
    "Q"=>"placeBirths",
    "R"=>"uStatus",
    "S"=>"docOriginal",
    "T"=>"docOriginalType",
    "U"=>"operator",
    "V"=>"address",
    "W"=>"addressActual",
    "X"=>"hostel",
    "Y"=>"mobile",
    "Z"=>"phone",
    "AA"=>"typeNP",
    "AB"=>"edName",
    "AC"=>"edType",
    "AD"=>"edCountry",
    "AE"=>"edRegion",
    "AF"=>"edLang",
    "AG"=>"edFinish",
    "AH"=>"edDocType",
    "AI"=>"edDocSerial",
    "AJ"=>"edDocNumber",
    "AK"=>"edDocDate",
    "AL"=>"edScore",
    "AQ"=>"docBenefits",
    "AV"=>"id",
    "AW"=>"id",
    "AZ"=>"appPriority",
    "BA"=>"appStatus",
    "BB"=>"specLevel",
    "BC"=>"specShape",
    "BD"=>"appCourse",
    "BE"=>"specCode",
    "BF"=>"specName",
    "BG"=>"appBasis",
    "BH"=>"numLD",
    "BI"=>"methodSubmitting",
    "BJ"=>"approval",
    "BK"=>"orderNum",
    "BL"=>"orderDate",
];
$trials= [
    "AR"=>"type",
    "AS"=>"discipline",
    "AT"=>"score",
    "AU"=>"date",
];
$familyLetters= [
    "AM"=>"fType",
    "AN"=>"fFIO",
    "AO"=>"fPhone",
    "AP"=>"fDoc",
];
$achvsLetters= [
    "AX"=>"achName",
    "AY"=>"achScore",
];

$sheet1->getStyle("A:BL")->applyFromArray($styles);
$sheet1->setAutoFilter('A2:BL2');
foreach ($lettersHead as $letter=>$field){
    $sheet1->setCellValueExplicit($letter."1",$field,\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
}

foreach ($res as $app){
    if(!empty($app->family))
        $app->family= json_decode($app->family);
    else
        $app->family= [];

    if(!empty($app->trials))
        $app->trials= json_decode($app->trials);
    else
        $app->trials= [];

    if(!empty($app->achs))
        $app->achs= json_decode($app->achs);
    else
        $app->achs= [];

    $rcf= count($app->family)?count($app->family):1;
    $rct= count($app->trials)?count($app->trials):1;
    $rca= count($app->achs)?count($app->achs):1;
    $rc= nok($rcf, $rct);
    $rc= nok($rc, $rca);

    $re= $rs+$rc-1;
    foreach ($lettersMax as $letter=>$field){
        $sheet1->setCellValueExplicit($letter.$rs,$app->{$field},\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
        $sheet1->mergeCells($letter.$rs.":".$letter.$re);
    }

    if(count($app->trials)){
        foreach ($app->trials as $key=>$trial){
            $coef= $rc/$rct;
            foreach ($trials as $letter=>$field){
                $sheet1->setCellValueExplicit($letter.($rs+$key*$coef),$trial->{$field},\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                if($coef!==1){
                    $merge= $letter.($rs+$key*$coef).":".$letter.($rs+$key*$coef+$coef-1);
                    $sheet1->mergeCells($merge);
                }
            }
        }
    }
    if(count($app->family)){
        foreach ($app->family as $key=>$family){
            $coef= $rc/$rcf;
            foreach ($familyLetters as $letter=>$field){
                $sheet1->setCellValueExplicit($letter.($rs+$key*$coef),$family->{$field},\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                if($coef>1){
                    $merge= $letter.($rs+$key*$coef).":".$letter.($rs+$key*$coef+$coef-1);
                    $sheet1->mergeCells($merge);
                }
            }
        }
    }

    if(count($app->achs)){
        foreach ($app->achs as $key=>$achs){
            $coef= $rc/$rca;
            foreach ($achvsLetters as $letter=>$field){
                $sheet1->setCellValueExplicit($letter.($rs+$key*$coef),$achs->{$field},\PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_STRING);
                if($coef>1){
                    $merge= $letter.($rs+$key*$coef).":".$letter.($rs+$key*$coef+$coef-1);
                    $sheet1->mergeCells($merge);
                }
            }
        }
    }

    $rs= $re+1;
}
/** SAVE */
$spreadsheet->setActiveSheetIndex(0);
$writer = new Xlsx($spreadsheet);
$writer->save("xls/report.xlsx");

echo "<a href='download.php' target='_blank'>Файл сформирован. Скачать.</a>";
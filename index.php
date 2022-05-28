<?php

ini_set('display_errors', 'off');
require_once('vendor/autoload.php');
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
$login = "usn_comp";
$password = "CoMpUtEx2021";
$ip = NULL;
function req($url)
{
    $ch = curl_init();
    curl_setopt($ch, CURLOPT_CONNECTTIMEOUT, 10);
    curl_setopt($ch, CURLOPT_TIMEOUT, 10);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_NOBODY, false);
    curl_setopt($ch, CURLOPT_HEADER, false);
    if (!is_null($ip)) {
        curl_setopt($ch, CURLOPT_INTERFACE, $ip);
    }
    curl_setopt($ch, CURLOPT_URL, $url);
    $res = curl_exec($ch);
    curl_close($ch);
    return $res;
}
function showResult($res)
{
    $obj = json_decode($res, true);
    if (strpos($obj['listStock']['rows'][0][8],' ') == FALSE)
        $BestPriceSupplier = $obj['listStock']['rows'][0][8];
        else $BestPriceSupplier = substr($obj['listStock']['rows'][0][8], 0, strpos($obj['listStock']['rows'][0][8],' '));
    $BestPrice = str_replace('~', '', $obj['listStock']['rows'][0][7]);
    $BestPrice = str_replace('p.', '', $BestPrice);
    $BestPrice = str_replace('р.', '', $BestPrice);
    $BestPrice = str_replace(',', '.', $BestPrice);

    if (strpos($obj['listStock']['rows'][1][8],' ') == FALSE)
        $SecondPriceSupplier = $obj['listStock']['rows'][1][8];
        else $SecondPriceSupplier = substr($obj['listStock']['rows'][1][8], 0, strpos($obj['listStock']['rows'][1][8],' '));
    $SecondPrice = str_replace('~', '', $obj['listStock']['rows'][1][7]);
    $SecondPrice = str_replace('p.', '', $SecondPrice);
    $SecondPrice = str_replace('р.', '', $SecondPrice);
    $SecondPrice = str_replace(',', '.', $SecondPrice);
    return [$BestPrice, $BestPriceSupplier, $SecondPrice, $SecondPriceSupplier];
}

echo '
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
    <meta http-equiv="Content-Style-Type" content="text/css">
    <title></title>
    <meta name="Generator" content="Cocoa HTML Writer">
    <meta name="CocoaVersion" content="1671.2">

    <style>
        @font-face {
            font-family: verdana; /* Имя шрифта */
            src: url(verdana.ttf); /* Путь к файлу со шрифтом */
            font-family: oldStile; /* Имя шрифта */
            src: url(MGS45__C.ttf); /* Путь к файлу со шрифтом */
        }
        body {
            font-family: verdana;
        }
        table {
            width: 80%;
            border-collapse: collapse;
        }
        td, th {
            border: 1px solid #f3f3f3;
            padding: 3px 7px 2px 7px;
        }
        tr {background-color: white;}
        tr:nth-child(odd) { background-color: rgb(243,243,243); }
        th {
            text-align: left;
            padding: 5px;
            background-color: rgb(233,239,245);
            color: #000000;
        }
    div.echo{
            font-family: oldStile;
            position: fixed;
            top: 0;
            right: 0;
            bottom: 0;
            left: 0;
            margin: auto;
            width:800px;
            height:100px;
            box-shadow: 0 0 30px lightslategrey;
            border-radius: 2px; 
            display: block;
            z-index:99999999;
            background-color: rgba(255,255,255,0.2);
            opacity: 0.8;
            color: black;
            font-weight: bold;
            font-size: 20px;
            padding: 15px;
            text-align: center;
        }
    div.debug{
            font-family: oldStile;
            position: fixed;
            top: 100px;
            right: 0;
            bottom: 0;
            left: 0;
            width:800px;
            height:400px;
            box-shadow: 0 0 30px lightslategrey;
            border-radius: 2px; 
            display: block;
            z-index:99999999;
            background-color: rgba(255,255,255,0.2);
            color: black;
            font-weight: bold;
            font-size: 20px;
            padding: 15px;
            text-align: center;
        }
    </style>
</head>';

if ($_POST['PercentOver']=='') $PercentOver = 6; else $PercentOver = $_POST['PercentOver'];
if ($_POST['PercentAlone']=='') $PercentAlone = 4; else $PercentAlone = $_POST['PercentAlone'];
if ($_POST['OurName']=='') $OurName = 'COMPUTEX'; else $OurName = $_POST['OurName'];
if ($_POST['Nacenka']=='') $Nacenka = 6; else $Nacenka = $_POST['Nacenka'];
if ($_POST['NDS']=='') $NDS = 20; else $NDS = $_POST['NDS'];
if ($_POST['Kurs']=='') $Kurs = 6.40; else $Kurs = $_POST['Kurs'];
if ($_POST['CutLower']=='') $CutLower = 1000; else $CutLower = $_POST['CutLower'];
if ($_POST['excelReadRowStart']=='') $excelReadRowStart = 8; else $excelReadRowStart = $_POST['excelReadRowStart'];


$excelReadRowCount = 0;
$rowCountForPercentDone = $excelReadRowStart;
$excelWriteRowCount = 7;
$excelWriteRowCount_D = 7;

echo "
    <form action='index.php' method='post' enctype='multipart/form-data' id='254' name='aaa'>
        <table style='font-size: larger;'>
            <tr>
                <td>
                    Если наша цена меньше ближайшей конкурентной цены на этот процент, то наша цена подтянется под конкурента, но останется меньше его цены на этот процент
                </td>
                <td> 
                    <input style='font-size: x-large;' value='$PercentOver' type='text' name='PercentOver' id='PercentOver' maxlength='2' size='2'> %
                </td>
            </tr>
            <tr>
                <td>
                    Если наше предложение единственное, то наша цена будет увеличена на этот дополнительный процент
                </td>
                <td>
                    <input style='font-size: x-large;' value='$PercentAlone' type='text' name='PercentAlone' id='PercentAlone' maxlength='2' size='2'> %
                </td>
            </tr>
            <tr>
                <td>
                    Название нашей компании в поисковой выдаче s4b
                </td>
                <td>
                    <input style='font-size: x-large;' value='$OurName' type='text' name='OurName' id='OurName' maxlength='15' size='15'>
                </td>
            </tr>
            <tr>
                <td>
                    Наценка в процентах на цену от входящего прайса 
                </td>
                <td>
                    <input style='font-size: x-large;' value='$Nacenka' type='text' name='Nacenka' id='Nacenka' maxlength='2' size='2'> %
                </td>
            </tr>
            <tr>
                <td>
                    Ставка НДС в РФ 
                </td>
                <td>
                    <input style='font-size: x-large;' value='$NDS' type='text' name='NDS' id='NDS' maxlength='2' size='2'> %
                </td>
            </tr>
            <tr>
                <td>
                    Курс рубля к тенге 
                </td>
                <td>
                    <input style='font-size: x-large;' value='$Kurs' type='text' name='Kurs' id='Kurs' maxlength='4' size='4'>
                </td>
            </tr>
            <tr>
                <td>
                    Не учитывать товары стоимость меньше
                </td>
                <td>
                    <input style='font-size: x-large;' value='$CutLower' type='text' name='CutLower' id='CutLower' maxlength='4' size='4'> руб.
                </td>
            </tr>
            <tr>
                <td>
                    Начинать анализ прайса со строки 
                </td>
                <td>
                    <input style='font-size: x-large;' value='$excelReadRowStart' type='text' name='excelReadRowStart' id='excelReadRowStart' maxlength='2' size='2'>
                </td>
            </tr>
            <tr>
                <td>
                    Публиковать товары с ценой хуже, чем другие два предложения 
                </td>
                <td>
                    <input style='font-size: x-large;' value='1' type='checkbox' name='PostOurExpencives' id='PostOurExpencives' checked >
                </td>
            </tr>
            <tr>
                <td>
                    Прайс моего поставщика 
                </td>
                <td>
                    <input style='font-size: large;' type='file' name='ioFileMy' id='ioFileMy' accept='.xlsx'>
                    <input style='font-size: large;' type='submit' value='Загрузить'>
                </td>
            </tr>
        </table>
    </form>
    <br>";
//echo '<br>=='.$_POST['$PostOurExpencives'].'==';
$PercentOver = $PercentOver/100+1;
$PercentAlone = $PercentAlone/100+1;
$Nacenka = $Nacenka/100+1;
$NDS = $NDS/100+1;

echo 'thios conmmit git text';

if ($_FILES['ioFileMy']['error'] == 0) {
//    echo 'Начало процесса в  <span style="color: red;">' . date("H:i:s",time()+21600) . '</span>';
    $ioFileMy = $_FILES['ioFileMy']['tmp_name'] . "_OAS.tmp";
    if (move_uploaded_file($_FILES['ioFileMy']['tmp_name'], $ioFileMy)) {
// для основного прайса для S4B
        $oSpreadsheet_Out = new Spreadsheet();
        $oSheet_Out = $oSpreadsheet_Out->getActiveSheet();
        $oWriter = IOFactory::createWriter($oSpreadsheet_Out, 'Xlsx');

        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A1', 'ComputeX / sborca.ru');
        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A2', 'Отдел продаж:  +7 495 727 33 53 sales@computex.ru');
        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A3', 'График работы: Пн-Пт с 10 до 19 часов');
        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A4', 'Москва. ул.Рабочая. 93с2');
        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A5', 'Цены для корпоративных клиентов от: '.date("j M Y",time()+10800).' '.date("G:i:s",time()+10800));
        $oSheet_Out->getStyle('A1:A5')->getFont()->setSize(16);
        $oSheet_Out->getStyle('A5')->getFont()->setBold(true);

        $oSheet_Out->getStyle('7')->getFont()->setSize(14);
        $oSheet_Out->getStyle('7')->getFont()->setBold(true);

        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A7', 'PartNumber');
        $oSheet_Out->getColumnDimension('A')->setWidth(22);

        $oSpreadsheet_Out->getActiveSheet()->setCellValue('B7', 'Наименование');
        $oSheet_Out->getColumnDimension('B')->setWidth(80);
        $oSheet_Out->getStyle('B')->getAlignment()->setWrapText(true);

        $oSpreadsheet_Out->getActiveSheet()->setCellValue('C7', 'Доступность');
        $oSheet_Out->getColumnDimension('C')->setWidth(14);

        $oSpreadsheet_Out->getActiveSheet()->setCellValue('D7', 'Цена руб.');
        $oSheet_Out->getColumnDimension('D')->setWidth(14);

//        $oSpreadsheet_Out->getActiveSheet()->setCellValue('E7', 'Гарантия');
//        $oSheet_Out->getColumnDimension('E')->setWidth(14);

        $oSpreadsheet_Out->getActiveSheet()->setCellValue('F7', 'Шт. в коробке');
        $oSheet_Out->getColumnDimension('F')->setWidth(14);
//// для основного прайса для S4B


// для прайса для Дистрибьюторов - без дополнительных наценок
        $oSpreadsheet_Out_D = new Spreadsheet();
        $oSheet_Out_D = $oSpreadsheet_Out_D->getActiveSheet();
        $oWriter_D = IOFactory::createWriter($oSpreadsheet_Out_D, 'Xlsx');

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('A1', 'ComputeX / sborca.ru');
        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('A2', 'Отдел продаж:  +7 495 727 33 53 sales@computex.ru');
        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('A3', 'График работы: Пн-Пт с 10 до 19 часов');
        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('A4', 'Москва. ул.Рабочая. 93с2');

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('A5', 'Специальные цены для Дистрибьюторов от: '.date("j M Y",time()+10800).' '.date("G:i:s",time()+10800).' Действуют при выкупе всей партии любой товарной позиции.');

        $oSheet_Out_D->getStyle('A1:A5')->getFont()->setSize(16);
        $oSheet_Out_D->getStyle('A5')->getFont()->setBold(true);

        $oSheet_Out_D->getStyle('7')->getFont()->setSize(14);
        $oSheet_Out_D->getStyle('7')->getFont()->setBold(true);

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('A7', 'PartNumber');
        $oSheet_Out_D->getColumnDimension('A')->setWidth(22);

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('B7', 'Наименование');
        $oSheet_Out_D->getColumnDimension('B')->setWidth(80);
        $oSheet_Out_D->getStyle('B')->getAlignment()->setWrapText(true);

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('C7', 'Доступность');
        $oSheet_Out_D->getColumnDimension('C')->setWidth(14);

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('D7', 'Цена руб.');
        $oSheet_Out_D->getColumnDimension('D')->setWidth(14);

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('E7', 'Ближайшая цена s4b>');
        $oSheet_Out_D->getColumnDimension('E')->setWidth(14);

        $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('F7', 'Шт. в коробке');
        $oSheet_Out_D->getColumnDimension('F')->setWidth(14);
//// для прайса для Дистрибьюторов - без дополнительных наценок



/*        class MyReadFilter implements PhpOffice\PhpSpreadsheet\Reader\IReadFilter
        {
            public function readCell($columnAddress, $row, $worksheetName = '')
            {
                // Read title row and rows 20 - 30
                if (
                    $columnAddress == 'B'
                    OR $columnAddress == 'C'
                    OR $columnAddress == 'D'
                    OR $columnAddress == 'E'
                    OR $columnAddress == 'F'
                    OR $columnAddress == 'G'
                    OR $columnAddress == 'A'
                    ) {
                    return true;
                }
                return false;
            }
        }
*/
        $oReader = new Xlsx();
        //$oReader->setReadFilter(new MyReadFilter());
        $oSpreadsheet = $oReader->load($ioFileMy);
        $oCells = $oSpreadsheet->getActiveSheet()->getCellCollection();
        $oRowEnd = $oCells->getHighestRow();
        $recountStartAt = time();
        echo '
            <table>
                <tr>
                    <th>#</th>
                    <th>Part Number</th>
                    <th>Наименование</th>
                    <th>Status</th>
                    <th>Моя цена</th>
                    <th>Лучшая цена</th>
                    <th>Поставщик</th>
                    <th>Вторая цена</th>
                    <th>Поставщик</th>
                    <th>В прайс</th>
                </tr>
            ';
        for ($iRow = $excelReadRowStart; $iRow <= $oRowEnd; $iRow++) {
        //for ($iRow = 31; $iRow <= 36; $iRow++) {
            $excelReadRowCount++;
            $rowCountForPercentDone++;
            //echo $excelReadRowCount . '. ';
            $MyPrice = str_replace(' ', '', $oCells->get('E' . $iRow));
            //echo '<br>'.$MyPrice.' '.$CutLower.'<br>';
//
           // <span style='vertical-align: center ; color: darkblue; font-weight: bold;'>

            //$recountStartAt = date("H:i:s",time()+21600);

            //$recountWillFinishAt = date("H:i:s",time()+21600+($oRowEnd*1.1));
            //$recountWillFinishAt = date("H : i",time()+21600+($oRowEnd - $excelReadRowCount)*(time() - $recountStartAt)/$excelReadRowCount);

            $recountWillFinishAt = date("i:s",($oRowEnd - $excelReadRowCount)*(time() - $recountStartAt)/$excelReadRowCount);

            echo "
                    <script>
                        document.getElementById('divEcho').remove();
                    </script>
                    <div class=echo id='divEcho'>Сравнение прайсов<br><br>
                            ".(round($rowCountForPercentDone*100/$oRowEnd,0))."% 
                            <img src='1pxblue.png' width='".(round($rowCountForPercentDone*100/$oRowEnd,0)*4)."px' height='20px'><img src='1pxgrey.png' width='".(round(100-( $rowCountForPercentDone*100/$oRowEnd),0)*4)."px' height='20px'>
                            $recountWillFinishAt мин<br>
                        $rowCountForPercentDone из $oRowEnd
                    </div>
                ";

            if (is_numeric($MyPrice) AND round($MyPrice/$Kurs*$Nacenka*$NDS,0) > $CutLower) {
                $MyPrice = round($MyPrice/$Kurs*$Nacenka*$NDS,0);
                //$MyPrice = round($MyPrice,0);
                $PartNumber = trim($oCells->get('B' . $iRow));

                if (strpos($PartNumber,'htt')!=FALSE){
                    $PartNumber = substr($PartNumber,0,strripos($PartNumber, '"'));
                    $PartNumber = substr($PartNumber,strripos($PartNumber, '"')+1);
                    $PartNumber = str_replace(',','.',$PartNumber);
                }

                $PartName = trim($oCells->get('C' . $iRow));
                $PartQuantity = trim($oCells->get('D' . $iRow));
                $PartGuarantee = trim($oCells->get('F' . $iRow));
                $PartQuantityInBox = trim($oCells->get('G' . $iRow));
                //echo $PartNumber.' = ' . $MyPrice;

                $query = urlencode($PartNumber);
                $url1 = 'http://s4b.ru/s.jsp?a=10041&at=3&usrLogin='.$login.'&usrPassword='.$password;
                $res1 = req($url1);
                $obj=json_decode($res1, true);
                if ($obj['status'] != 'ok') {
                    exit('error:'.$obj['status']);
                }
                $urlBase = $obj['url'];



//                while ($res2[0]==> string(0) "" [1]=> NULL [2]=> string(0) "" [3]=> NULL)





                $res2 = req($urlBase.'&a=6001&sr='.$query);
                $QueryResult = showResult($res2);



                $BestPrice = round($QueryResult[0],0);
                $BestPriceSupplier = $QueryResult[1];
                $SecondPrice = round($QueryResult[2],0);
                $SecondPriceSupplier = $QueryResult[3];

//                if ( $SecondPrice != 0 AND $MyPrice > $SecondPrice ) $HaveToWrite = FALSE; ELSE $HaveToWrite = TRUE;


                        //echo ' best='.$BestPrice . ' by '. $BestPriceSupplier. ' &nbsp &nbsp &nbsp &nbsp &nbsp second='. $SecondPrice . ' by '.$SecondPriceSupplier;

                        //if ($BestPriceSupplier == $OurName)
                        //    $NotMyPrice = $SecondPrice;
                        //else $NotMyPrice = $BestPrice;
                        //echo ' => NotMyPrice='.$NotMyPrice.'<br>';
                        //$MyPrice = round($MyPrice, 0);
                        //$NotMyPrice = 0;
                        $Status = 0;
                        $styleForPrice = '';
                        //echo '<br>BestPriceSupplier='.$BestPriceSupplier.'/ OurName='.$OurName.' '.(($BestPriceSupplier == $OurName)?'TRUE':'FALSE').'/<br>';
                        if (($BestPriceSupplier == $OurName OR $BestPrice == 0) AND $SecondPrice == 0){
                            $Status = 1 ; //Мы единственные
                            $styleForPrice = ' color: RoyalBlue; ';
                            $PriceToPrice = round($MyPrice * $PercentAlone,0);
                        }
                        if ($Status == 0 AND ($MyPrice <= $BestPrice AND ($BestPriceSupplier != $OurName OR ($BestPriceSupplier == $OurName AND $SecondPriceSupplier != NULL AND $SecondPriceSupplier!=$OurName)))){
                            $Status = 2 ; //Мы лучшие
                            $styleForPrice = ' color: darkgreen; ';
                            $PriceToPrice = $MyPrice;

                            if ($PartNumber=='LS27A600NWIXCI')
                                echo'';

                            if (($MyPrice <= round($BestPrice / $PercentOver,0)) AND $SecondPrice<>0) {
                                if ($BestPriceSupplier != $OurName )
                                    $PriceToPrice = round($BestPrice / $PercentOver,0);
                                else
                                    $PriceToPrice = round($SecondPrice / $PercentOver,0);

                                    $Status =3; //Мы лучшие и ниже рынка - нужно скорректировать цену
                                    $styleForPrice = ' color: darkgreen; font-weight:bold; ';
                                    $excelWriteRowCount_D++;


                                if ($excelWriteRowCount_D/2 == round($excelWriteRowCount_D/2)){
                                    $oSheet_Out_D->getStyle('A' . $excelWriteRowCount_D.':F'.$excelWriteRowCount_D)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
                                    $oSheet_Out_D->getStyle('A' . $excelWriteRowCount_D.':F'.$excelWriteRowCount_D)->getFill()->getStartColor()->setRGB('F3F9FF');
                                }

                                    $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('A' . $excelWriteRowCount_D, $PartNumber);
                                    $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('B' . $excelWriteRowCount_D, $PartName);
                                    $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('C' . $excelWriteRowCount_D, $PartQuantity);
                                    $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('D' . $excelWriteRowCount_D, $MyPrice);
                                    $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('E' . $excelWriteRowCount_D, $SecondPrice);
                                    $oSpreadsheet_Out_D->getActiveSheet()->setCellValue('F' . $excelWriteRowCount_D, $PartQuantityInBox);
                            }
                        }
                        if ($Status == 0 AND ($MyPrice > $BestPrice AND $BestPrice!=0 AND ($MyPrice <= $SecondPrice OR $SecondPrice == 0))){
                            $Status = 4;//Мы вторые
                            $styleForPrice = ' color: DarkGoldenRod; ';
                            $PriceToPrice = $MyPrice;
                        }
                        if ($Status == 0 AND ($MyPrice > $SecondPrice AND $SecondPrice > 0 )) {
                            $Status = 5;//Мы третьи или дальше.
                            $styleForPrice = ' color: darkred; ';
                            $PriceToPrice=$MyPrice;
                        }
                        $OurPosition[$Status]++;
                        echo '
                            <tr style="background-color: '.(($rowCountForPercentDone/2 == round($rowCountForPercentDone/2))?'rgb(243,249,255)':'white').';'.$styleForPrice.'">';
                        echo '
                                <td>'.$excelReadRowCount.'</td>
                                <td>'.$PartNumber.'</td>
                                <td>'.$PartName.'</td>
                                <td>'.$Status.'</td>
                                <td>'.$MyPrice.'</td>';
                        //$PriceToPrice = $MyPrice;
                        //echo '$MyPrice='.$MyPrice.' $BestPrice='.$BestPrice. ' $SecondPrice='.$SecondPrice;

                        //if ($Status == 0) $PriceToPrice = round($MyPrice * $PercentAlone,0);
                        //echo ' $MyPrice * $PercentAlone='.$MyPrice.'<br>';
                        echo '
                                <td>'.(($BestPriceSupplier == $OurName OR $BestPrice==0)?'':$BestPrice).'</td>
                                <td>'.(($BestPriceSupplier == $OurName)?'':$BestPriceSupplier).'</td>
                                <td>'.(($SecondPriceSupplier == $OurName OR $SecondPrice==0)?'':$SecondPrice).'</td>
                                <td>'.(($SecondPriceSupplier == $OurName)?'':$SecondPriceSupplier).'</td>
                                <td>'.$PriceToPrice.'</td>
                            </tr>';
                    if ($Status<5 OR isset($_POST["PostOurExpencives"])==TRUE) {
                        $excelWriteRowCount++;
                        if ($excelWriteRowCount/2 == round($excelWriteRowCount/2)){
                            $oSheet_Out->getStyle('A' . $excelWriteRowCount.':F'.$excelWriteRowCount)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
                            $oSheet_Out->getStyle('A' . $excelWriteRowCount.':F'.$excelWriteRowCount)->getFill()->getStartColor()->setRGB('F3F9FF');
                        }
                        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A' . $excelWriteRowCount, $PartNumber);
                        $oSpreadsheet_Out->getActiveSheet()->setCellValue('B' . $excelWriteRowCount, $PartName);
                        $oSpreadsheet_Out->getActiveSheet()->setCellValue('C' . $excelWriteRowCount, $PartQuantity);
                        $oSpreadsheet_Out->getActiveSheet()->setCellValue('D' . $excelWriteRowCount, $PriceToPrice);
                        $oSpreadsheet_Out->getActiveSheet()->setCellValue('E' . $excelWriteRowCount, $PartGuarantee);
                        $oSpreadsheet_Out->getActiveSheet()->setCellValue('F' . $excelWriteRowCount, $PartQuantityInBox);
                    }

                }
                else {
                    if (!is_numeric($MyPrice)){
                        echo '
                        <tr>
                            <td colspan=10 style="text-align: center; font-size: larger; background-color: royalblue; color: lightgrey;">'.
                            $oCells->get('A' . $iRow).
                            '</td>
                        </tr>';
                        $excelReadRowCount--;
                        $excelWriteRowCount++;

                        $oSheet_Out->getStyle('A' . $excelWriteRowCount)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
                        $oSheet_Out->getStyle('A' . $excelWriteRowCount)->getFill()->getStartColor()->setRGB('DFE5EB');

                        $oSpreadsheet_Out->getActiveSheet()->setCellValue('A' . $excelWriteRowCount, $oCells->get('A' . $iRow));
                        $oSpreadsheet_Out->getActiveSheet()->mergeCells('A'.$excelWriteRowCount.':F'.$excelWriteRowCount);
                        $oSpreadsheet_Out->getActiveSheet()->getRowDimension($excelWriteRowCount)->setRowHeight(30);
                        $oSpreadsheet_Out->getActiveSheet()->getStyle('A' . $excelWriteRowCount)->getFont()->setSize(16);
                        $oSpreadsheet_Out->getActiveSheet()->getStyle('A' . $excelWriteRowCount)->getFont()->setBold(true);
                        //$oSpreadsheet_Out->getActiveSheet()->getStyle('A' . $excelWriteRowCount)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
                        //$oSpreadsheet_Out->getActiveSheet()->getStyle('A' . $excelWriteRowCount)->getFont()->getColor()->setRGB('EEEEEE');
                        //->getStyle()->getFont()->getColor()->setARGB(
                    }
                }
                sleep(1);
            }
        }
        echo '</table>';
        $oWriter->save('price.xlsx');
        $oWriter_D->save('price_D.xlsx');
    }
echo "
                    <script>
                        var el = document.getElementById('divEcho');
                        el.remove();
                    </script>
                    <div class=echo id='divEcho' style='background-color: rgba(255,255,255,0.6);'>Сравнение прайсов завершено<br><br>
                            $oRowEnd строк<br>
                            за ".date("i мин s сек",time()-$recountStartAt)."
                    </div>
                    <script>
                        var el = document.getElementById('divEcho');
                        el.onclick = function() {
                        el.remove();}
                    </script>
                ";

echo '<br>...process end at <span style="color: red;">'.date("H:i:s",time()+21600).'</span><br><br>';
Echo
    'Из <span style="font-weight: bold; font-size: larger;">'.$excelReadRowCount.'</span> товаров с ценой больше <span style="font-weight: bold; font-size: larger;">'.$CutLower.'</span> руб. у нас лучшая цена по <span style="font-weight: bold; color: darkolivegreen; font-size: larger;">'.($OurPosition[1]+$OurPosition[2]+$OurPosition[3]).'</span> позициям, из них:<br>
     <span style="color: RoyalBlue">&nbsp - Товар есть только у нас: <span style="font-size: larger;">'.$OurPosition[1].'</span></span><br>
     <span style="color: DarkGreen; font-weight:bold">&nbsp - Наша цена значительно ниже: <span style="font-size: larger;">'.$OurPosition[3].'</span></span><br>
     <span style="color: darkgreen">&nbsp - Наша цена первая: <span style="font-size: larger;">'.$OurPosition[2].'</span></span><br>
     <span style="color: DarkGoldenRod">&nbsp - Наша цена вторая: <span style="font-size: larger;">'.$OurPosition[4].'</span></span><br>
     <span style="color: darkred">&nbsp - Наша цена третья или хуже: <span style="font-size: larger;">'.$OurPosition[5].'</span></span>';



/*
array(4) { [0]=> string(5) "69206" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "56588" [1]=> string(8) "Pleer.ru" [2]=> string(5) "57259" [3]=> string(6) "NetLab" }
array(4) { [0]=> string(5) "55587" [1]=> string(4) "Elko" [2]=> string(5) "68336" [3]=> string(8) "COMPUTEX" }
array(4) { [0]=> string(5) "55404" [1]=> string(6) "NetLab" [2]=> string(8) "58202.53" [3]=> string(8) "TechnoIt" }
array(4) { [0]=> string(5) "72897" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "17164" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "13380" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "20023" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "24638" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "23808" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "20208" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(4) "9636" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(4) "2995" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(4) "3086" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(8) "37721.60" [1]=> string(12) "MobiСomShop" [2]=> string(8) "38000.82" [3]=> string(8) "TechnoIt" }
array(4) { [0]=> string(5) "34999" [1]=> string(3) "DNS" [2]=> string(5) "42656" [3]=> string(8) "COMPUTEX" }
array(4) { [0]=> string(5) "27607" [1]=> string(8) "Pleer.ru" [2]=> string(8) "28547.38" [3]=> string(7) "ReStart" }
array(4) { [0]=> string(5) "40835" [1]=> string(8) "Pleer.ru" [2]=> string(5) "46225" [3]=> string(8) "COMPUTEX" }
array(4) { [0]=> string(5) "42830" [1]=> string(8) "COMPUTEX" [2]=> string(5) "44584" [3]=> string(8) "Pleer.ru" }
array(4) { [0]=> string(5) "14118" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(5) "17532" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }
array(4) { [0]=> string(4) "2122" [1]=> string(8) "COMPUTEX" [2]=> string(0) "" [3]=> NULL }


 */
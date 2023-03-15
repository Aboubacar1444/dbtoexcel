<?php

require "vendor/autoload.php";


try {
    $connexion = new PDO('mysql:host=localhost;dbname=samamoney;charset=utf8mb4', 'root', 'root');
} catch (Exception $e) {
    die('error:' . $e->getMessage());
}

//initialize date for daily report
$date = date('Y-m-d 00:00:00', strtotime('yesterday'));
$endDate = date('Y-m-d 23:59:59', strtotime('yesterday'));

//Req to table sama_marchand_histo
$req = "SELECT * FROM sama_marchand_histo     WHERE date_remonter BETWEEN '$date' AND '$endDate'";
$pr = $connexion->prepare($req);
$pr->execute();

//Req to table sama_modena_remonter
$req1 = "SELECT * FROM sama_modena_remonter_uv WHERE date_remonter BETWEEN '$date' AND '$endDate' ";
$pr1 = $connexion->prepare($req1);
$pr1->execute();







$spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

//Initialize row number
$sn = 1;

//Header Table sama_modena_remonter_uv
$sheet->setCellValue('A'.$sn,'Date');
$sheet->setCellValue('B'.$sn,'Heure');
$sheet->setCellValue('C'.$sn,'Référence');
$sheet->setCellValue('D'.$sn,'Compte marchand principal');
$sheet->setCellValue('E'.$sn,'Compte marchand');
$sheet->setCellValue('F'.$sn,'Montant');
$sheet->setCellValue('G'.$sn,'Status');

//Fetch  Table sama_modena_remonter_uv values
while ($row = $pr1->fetch()) {
    $sn++;
    $formatDate             = substr($row["date_remonter"], 0, 10);
    $formatDate             = date_format( new \Datetime($formatDate), "d/m/Y");
    $formatTime             = substr($row["date_remonter"], 10);
    $transacNumber          = $row["transac_number"];
    $marchandReception      = $row["marchand_reception"];
    $codeMarchand           = $row["code_marchand"];
    $montant                = $row["montant"];
    $status                 = "Success";



    $sheet->setCellValue("A".$sn, $formatDate);
    $sheet->setCellValue("B".$sn, $formatTime);
    $sheet->setCellValue("C".$sn, $transacNumber);
    $sheet->setCellValue("D".$sn, $marchandReception);
    $sheet->setCellValue("E".$sn, $codeMarchand);
    $sheet->setCellValue("F".$sn, $montant);
    $sheet->setCellValue("G".$sn, $status);

}
//_____________________________END First Part____________________________

//Create space between the tables
$sn+= 5;

//Header Table sama_marchand_histo
$sheet->setCellValue('A'.$sn,'Date');
$sheet->setCellValue('B'.$sn,'Heure');
$sheet->setCellValue('C'.$sn,'Ref Commande');
$sheet->setCellValue('D'.$sn,'Ref Achat SAMA');
$sheet->setCellValue('E'.$sn,'Nom marchand');
$sheet->setCellValue('F'.$sn,'Compte Marchand');
$sheet->setCellValue('G'.$sn,'Montant');
$sheet->setCellValue('H'.$sn,'Status');


//Fetch  Table sama_marchand_histo values
while ($row = $pr->fetch()) {
    $sn++;
    $formatDate             = substr($row["dateEnreg"], 0, 10);
    $formatDate             = date_format( new \Datetime($formatDate), "d/m/Y");

    $formatTime             = substr($row["dateEnreg"], 10);
    $idCommande             = $row["idCommande"];
    $marchandReception      = $row["transNumber"];
    $nomMarchand            = $row["nomMarchand"];
    $codeMarchand           = $row["cmd"];
    $montant                = $row["montant"];
    $status                 = $row["msg"] == "Echec::Mot de passe incorrect, veuillez le ressais" ? "Waiting": $row["msg"]  ;



    $sheet->setCellValue("A".$sn, $formatDate);
    $sheet->setCellValue("B".$sn, $formatTime);
    $sheet->setCellValue("C".$sn, $idCommande);
    $sheet->setCellValue("D".$sn, $marchandReception);
    $sheet->setCellValue("E".$sn, $nomMarchand);
    $sheet->setCellValue("F".$sn, $codeMarchand);
    $sheet->setCellValue("G".$sn, $montant);
    $sheet->setCellValue("H".$sn, $status);

}




//Initialized writer and save file
$filename = "SAMAMONEY_DB_To_EXCEL_" . date('d_m_Y_H_i_s', strtotime('yesterday'));
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
$writer->save("generate/$filename".'.xlsx');

//Mail protocole

try {

//    use PHPMailer\PHPMailer\SMTP;
    $mail = new PHPMailer\PHPMailer\PHPMailer(true);
    //$mail->SMTPDebug = SMTP::DEBUG_SERVER;
    $mail->isSMTP();
    $mail->isHTML();
    $mail->SMTPSecure = \PHPMailer\PHPMailer\PHPMailer::ENCRYPTION_SMTPS;
    $mail->Host = 'smtp.hostinger.com';
    $mail->Port = 465;
    $mail->SMTPAuth = true;
    $mail->Username = 'xyz@wenovate-ml.com';
    $mail->Password = 'Wenovate@2022';
    $mail->setFrom('xyz@wenovate-ml.com', 'no-reply[dailyMail]');
    $mail->addAddress('kane@sama.money', 'Mr. Kane');
    $mail->Subject = 'daily Excel file';
    $mail->addAttachment("./generate/$filename");
    $mail->msgHTML(file_get_contents('message.html'), __DIR__, true);
    $mail->Body = file_get_contents('message.html');
    $mail->send();
    echo 'Message has been sent';
} catch (\Throwable $th) {
    throw $th;
}

echo "OK";





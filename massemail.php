<?php
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

require 'vendor/autoload.php';

$emailList = array();
$messageSent = false;
$errorMsg = '';

if ($_SERVER['REQUEST_METHOD'] == 'POST' && $_FILES['file']['error'] == 0) {
    $spreadsheet = IOFactory::load($_FILES['file']['tmp_name']);
    $columnIndex = Coordinate::columnIndexFromString('A');
    $highestRow = $spreadsheet->getActiveSheet()->getHighestRow();

    for ($row = 1; $row <= $highestRow; ++$row) {
        $email = $spreadsheet->getActiveSheet()->getCellByColumnAndRow($columnIndex, $row)->getValue();
        if (filter_var($email, FILTER_VALIDATE_EMAIL)) {
            $emailList[] = $email;
        }
    }

    $mail = new PHPMailer(true);

    try {
        $mail->isSMTP();
        $mail->SMTPAuth = true;
        $mail->Host = 'smtp-mail.outlook.com';
        $mail->Username = 'iheb.chebbi@lcieducation.net';
        $mail->Password = 'Azerty123';
        $mail->Port = 587;
        $mail->SMTPSecure = 'tls';

        $mail->setFrom('iheb.chebbi@lcieducation.net', 'Votre Nom');

        foreach ($emailList as $recipient) {
            $mail->addAddress($recipient);
        }

        $message = isset($_POST['message']) ? $_POST['message'] : '';
        $subject = isset($_POST['subject']) ? $_POST['subject'] : '';

        $mail->Subject = $subject;
        $mail->Body = '<html>
                        <body style="font-family: Arial, sans-serif; background-color: #f5f5f5; padding: 20px;">
                            <h2 style="color: #007bff;">Bonjour !</h2>
                            <p>Votre contenu personnalisé de l\'e-mail va ici.</p>
                            <p>Votre message : ' . nl2br($message) . '</p>
                            <p>Cordialement,<br>Votre Nom</p>
                        </body>
                      </html>';
        $mail->AltBody = 'Votre contenu personnalisé de l\'e-mail va ici. Votre message : ' . $message . ' Cordialement, Votre Nom';

        $mail->isHTML(true);

        if ($mail->send()) {
            $messageSent = true;
        } else {
            $errorMsg = 'L\'e-mail n\'a pas pu être envoyé. Erreur du serveur : ' . $mail->ErrorInfo;
        }
    } catch (Exception $e) {
        $errorMsg = 'L\'e-mail n\'a pas pu être envoyé. Erreur du serveur : ' . $mail->ErrorInfo;
    }
}
?>

<!DOCTYPE html>
<html lang="fr">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Envoi d'e-mail</title>
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/css/bootstrap.min.css">
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f5f5f5;
        }

        .container {
            margin-top: 50px;
        }

        .card {
            border: none;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }

        .card-header {
            background-color: #007bff;
            color: #fff;
            text-align: center;
            border-bottom: none;
        }

        .card-body {
            padding: 20px;
        }

        .btn-primary {
            background-color: #007bff;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }

        .btn-primary:hover {
            background-color: #0056b3;
        }

        .success {
            color: green;
        }

        .error {
            color: red;
        }
    </style>
</head>

<body>
    <div class="container">
        <div class="card">
            <div class="card-header">
                <h5>MAILEUR V1.0</h5>
            </div>
            <div class="card-body">
                <form method="post" action="" enctype="multipart/form-data">
                    <div class="form-group">
                        <label for="file">Télécharger la liste des emails :</label>
                        <input type="file" class="form-control-file" name="file" id="file" accept=".xls, .xlsx" required>
                    </div>
                    <div class="form-group">
                        <label for="subject">Sujet :</label>
                        <input type="text" class="form-control" name="subject" id="subject" required>
                    </div>
                    <div class="form-group">
                        <label for="message">Message :</label>
                        <textarea class="form-control" name="message" id="message" rows="4" required></textarea>
                    </div>
                    <button type="submit" class="btn btn-primary">Envoyer l'e-mail</button>
                </form>

                <?php if ($messageSent): ?>
                    <h2 class="success">E-mail envoyé avec succès à <?php echo count($emailList); ?> destinataires.</h2>
                <?php elseif ($errorMsg): ?>
                    <h2 class="error"><?php echo $errorMsg; ?></h2>
                <?php endif; ?>

                <?php if (!empty($emailList)): ?>
                    <h2>Adresses e-mail extraites :</h2>
                    <ul>
                        <?php foreach ($emailList as $email): ?>
                            <li><?php echo $email; ?></li>
                        <?php endforeach; ?>
                    </ul>
                <?php endif; ?>
            </div>
        </div>
    </div>

    <script src="https://code.jquery.com/jquery-3.3.1.slim.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.14.7/umd/popper.min.js"></script>
    <script src="https://stackpath.bootstrapcdn.com/bootstrap/4.3.1/js/bootstrap.min.js"></script>
</body>

</html>

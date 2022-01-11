<?php
set_time_limit(0);

require __DIR__ . '/vendor/autoload.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;

if (!file_exists(__DIR__ . '/email.md'))
{
  die("Please copy email.dist.md into email.md and adjust as necessary.");
}

$excelFiles = glob(__DIR__ . '/excel/*.csv');

if (empty($excelFiles)) {
  die("Please copy the Excel file with all emails in the excel folder.");
}

// false : no email address was POSTed
// true : an email address was posted, found and an email was sent to the parents
// A string : the error to show
$bEmailSent = false;

if (!empty($_POST['email'])) {

  $emailToRecover = trim(strtolower($_POST['email']));

  $foundRows = array();
  $bEmailSent = "This email address ($emailToRecover) was not found. Please use one of the PARENT's email address.
                 If you don't remember the email address used for the Aseica registration, please
                 send an email to <a href=\"mailto:google_help@aseica.org\">google_help@aseica.org</a>
                 with your full name and we'll try to sort it out!";

  foreach ($excelFiles as $excelFilename) {
    // $reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader('Xlsx');
    // $reader->setReadDataOnly(true);
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();

    $reader->setInputEncoding('UTF-8');
    $reader->setDelimiter(',');
    $reader->setEnclosure('');
    $reader->setSheetIndex(0);

    $spreadsheet = $reader->load($excelFilename);
    $worksheet = $spreadsheet->getActiveSheet();
  //  $worksheet = $spreadsheet->getSheetByName('CSV Pour export');

    $res = array();

    foreach ($worksheet->getRowIterator() as $row) {
      $cellIterator = $row->getCellIterator();
      $cellIterator->setIterateOnlyExistingCells(FALSE);

      $row = array();
      foreach ($cellIterator as $k => $cell) {
        $row[] = $cell->getCalculatedValue();
      }

      list($firstName, $lastName, $emailAddress, $password, $passwordHashFunction, $orgUnitPath, $newPrimaryEmail, $recoveryEmail, $homeSecondaryEmail) = $row;

      if ((strpos($recoveryEmail, '@') !== false && trim(strtolower($recoveryEmail)) == $emailToRecover) ||
          (strpos($homeSecondaryEmail, '@') !== false && trim(strtolower($homeSecondaryEmail)) == $emailToRecover)) {
        $foundRows[$emailAddress] = array('firstName' => $firstName, 'lastName' => $lastName, 'emailAddress' => $emailAddress, 'password' => $password, 'recoveryEmail' => $recoveryEmail, 'homeSecondaryEmail' => $homeSecondaryEmail);
      }
    }

    if (!empty($foundRows))
       $bEmailSent = sendEmail($foundRows);
  }
}


function sendEmail($rows)
{
  $parser = new \cebe\markdown\GithubMarkdown();
  $parser->html5 = true;
  $markdown = file_get_contents(__DIR__ . '/email.md');
  $htmltemplate = $parser->parse($markdown);

  $GLOBALS['mail'] = new PHPMailer(true);
  $GLOBALS['mail']->CharSet = "UTF-8";
  $GLOBALS['mail']->setFrom('google_help@aseica.org', 'Aseica Google Help');

  if (file_exists(__DIR__ . 'smtp_settings.php'))
    include_once(__DIR__ . 'smtp_settings.php');

  $recipients = array();

  foreach ($rows as $row) {
    if (!empty($row['recoveryEmail']))
      $recipients[$row['recoveryEmail']] = $row['recoveryEmail'];

    if (!empty($row['homeSecondaryEmail']))
      $recipients[$row['homeSecondaryEmail']] = $row['homeSecondaryEmail'];
  }

  //Recipients
  foreach ($recipients as $aRecipient)
    $GLOBALS['mail']->addAddress($aRecipient);

  $GLOBALS['mail']->isHTML(true); //Set email format to HTML

  foreach ($rows as $row) {

    $html = str_replace('LASTNAME', $row['lastName'], $htmltemplate);
    $html = str_replace('FIRSTNAME', $row['firstName'], $html);
    $html = str_replace('EMAIL', $row['emailAddress'], $html);
    $html = str_replace('PASSWORD', $row['password'], $html);

    try {
      $GLOBALS['mail']->Subject = 'Aseica email account information - ' . $row['firstName'] . ' ' . $row['lastName'];
      $GLOBALS['mail']->Body = $html;

      $GLOBALS['mail']->send();
    } catch (Exception $e) {
      return "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
    }
  }

  return true;
}

?>
<!doctype html>
<html lang="en">

<head>
  <!-- Required meta tags -->
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">

  <!-- Bootstrap CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-1BmE4kWBq78iYhFldvKuhfTAU6auU8tT94WrHftjDbrCEXSU1oBoqyl2QvZ6jIW3" crossorigin="anonymous">
  <link href="css/styles.css" rel="stylesheet">

  <title>Aseica account retrieval</title>
</head>

<body>

  <div class="container">
    <div class="row align-items-start">
      <div class="col-lg-5 mx-auto">
        <h1>Aseica account retrieval</h1>

        <form method="POST" class="mb-3">
          <div class="mb-3">
            <label for="exampleInputEmail1" class="form-label">Parental email address:</label>
            <input name="email" type="email" class="form-control"  placeholder="parent.email@example.com" id="exampleInputEmail1" aria-describedby="emailHelp">
            <div id="emailHelp" class="form-text">Please enter the email address of one of the parent, as given to the Aseica during registration.</div>
          </div>
          <button type="submit" class="btn btn-primary">Submit</button>
        </form>


<?php

if ($bEmailSent !== false)
{
  if ($bEmailSent === true)
  {
    echo "<div class=\"alert alert-success\" role=\"alert\">
    An email has been sent to this email address. Please check your email for the instructions on how to activate your Aseica account.
  </div>";
  }
  else
  {
    echo "<div class=\"alert alert-danger\" role=\"alert\">$bEmailSent</div>";
  }

}
?>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>

</body>

</html>
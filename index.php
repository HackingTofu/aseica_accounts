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
  $GLOBALS['mail']->Priority = 3;

  if (file_exists(__DIR__ . 'dkim.private'))
  {
    $GLOBALS['mail']->DKIM_domain = 'aseica.org';
    $GLOBALS['mail']->DKIM_private = __DIR__ . 'dkim.private';
    $GLOBALS['mail']->DKIM_selector = 'google';
    $GLOBALS['mail']->DKIM_passphrase = '';
    $GLOBALS['mail']->DKIM_identity = 'google_help@aseica.org';
  }

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


  <!-- Global site tag (gtag.js) - Google Analytics -->
  <script async src="https://www.googletagmanager.com/gtag/js?id=UA-241152566-1"></script>
  <script>
    window.dataLayer = window.dataLayer || [];
    function gtag(){dataLayer.push(arguments);}
    gtag('js', new Date());

    gtag('config', 'UA-241152566-1');
  </script>

</head>

<body>

  <div class="container">
    <div class="row align-items-start">
      <div class="col-lg-10 mx-auto">
        <div class="text-center">
          <img class="my-5" src="images/image1.png">
          <h1 class="mb-5" >Aseica account retrieval</h1>
        </div>

    <p>Welcome to the Aseica Google account page. Here, you&#39;ll be able to retrieve your
            Google Account for all class activities (Google Classroom, Google Meet, Google docs, etc.).</p>
    <p></p>
    <p></p>
    <p>The first step is to get your student email address and your initial password. Once
            you have it, you will be immediately prompted to change your password. Once your password has been changed,
            and you forget it, you can retrieve it easily by checking the procedure below.</p>
    <p></p>
    <p></p>
    <p>If you tried the steps in the tutorial and are still encountering difficulties, or
            if you have not received the message about the account, please contact <a href="mailto:google_help@aseica.org">google_help@aseica.org</a>.</p>
    <p></p>
    <p></p>
    <p>Table of contents</p>
    <ul>
        <li><a href="#retrieve">Getting your student email address</li>
        <li><a href="#login">Login in for the first time</a></li>
        <li><a href="#waystoemail">Ways to check your email</a></li>
        <li><a href="#redirect">Creating a redirect</a></li>
        <li><a href="#lostpassword">Retrieving your lost password</a></li>
    </ul>
    <p></p>
    <p></p>
    <hr>
    <p></p>
    <p></p>


    <h2 id="retrieve">Getting your student email address</h2>
    <p>In order to get your student email address, please enter below the address of one of
            the parent, as given to the Aseica during registration:</p>
    <p></p>

    <form method="POST" class="mb-3">
          <div class="mb-3">
            <label for="exampleInputEmail1" class="form-label">Parental email address:</label>
            <input name="email" type="email" class="form-control"  placeholder="parent.email@example.com" id="exampleInputEmail1" aria-describedby="emailHelp">
            <div id="emailHelp" class="form-text">Please enter the email address of one of the parent, as given to the Aseica during registration.</div>
          </div>
          <button type="submit" class="btn btn-primary">Send my email address to my parents now !</button>
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
    <p></p>
    <p>After clicking the button, you will receive an email at this address with
            instructions to log in for all Aseica registered kids in the family.</p>


    <h2 id="login">Login in for the first time</h2>
    <p><b>Step 1:</b> go to <a href="https://mail.google.com" target="_blank">https://mail.google.com</a></p>
    <p></p>
    <p>Do you see a window like the screenshot below?</p>
    <p></p>
    <p><img src="images/image6.png" style="max-width:100%"></p>
    <p></p>
    <ul>
        <li>NO &rarr; go to step 2</li>
        <li>YES &rarr; go to step 3 (skipping step 2)</li>
    </ul>
    <p></p>
    <p><b>Step 2:</b> Add another Google account</p>
    <p>(not needed unless you get a screen like below)</p>
    <p></p>
    <p>You&rsquo;re getting this if you are already logged in to one or more Google/Gmail accounts on the browser. You need to add the ASEICA account, so that you can switch easily between all Gmail-based accounts.</p>
    <p></p>
    <p>1/ Click on your profile picture/logo at the top right of the screen</p>
    <p><img src="images/image13.png" style="max-width:100%"></p>
    <p></p>
    <p></p>
    <p>2/ In the window that appears, select &ldquo;Add another account&rdquo;</p>
    <p><img src="images/image3.png" style="max-width:100%"></p>
    <p></p>
    <p></p>
    <p><b>Step 3:</b> enter login (student.XYZ@aseica.org) and &ldquo;Next&rdquo; </p>
    <p></p>
    <p>Follow the flow of the following pages (password, accept terms, change password, etc). You&rsquo;re done!</p>
    <p></p>
    <p>If you get screens not mentioned in this tutorial, please email a screenshot to google_help@aseica.org, with explanations of the steps you took. Thanks!</p>
    <h2 id="waystoemail">Ways to check your email</h2>
    <h4>Go to <a href="https://mail.google.com" target="_blank">mail.google.com</a></h4>
    <p>The simplest way to consult this email is to go to mail.google.com</p>
    <p></p>
    <h4>Forward to an existing address</h4>
    <p>You can forward all email to an existing address. See <a href="#redirect">further</a>.</p>
    <p></p>
    <h4>Use Outlook or an email client</h4>
    <p>You can use <a href="https://support.google.com/mail/answer/7104828?hl=en-GB&visit_id=637786964639990659-662828657&rd=1" target="_blank">imap or pop</a> to download this email to any email client such as Outlook.</p>
    <p></p>
    <h4>Add this account to a smartphone</h4>
    <p>You can add the aseica email address to a smartphone.</p>



    <h2 id="redirect">Creating a redirect</h2>
    <p>You can decide to forward your email to an existing account your child already checks regularly.</p>
    <p></p>
    <p><b>Step 1:</b> go to your Aseica email settings</p>
    <p>1/ Click on the cog icon</p>
    <p><img src="images/image4.png" style="max-width:100%"></p>
    <p>2/ Click on &ldquo;See all settings&rdquo;</p>
    <p><img src="images/image7.png" style="max-width:100%"></p>
    <p></p>
    <p><b>Step 2:</b> Select the &ldquo;Forwarding and
            POP/IMAP&rdquo; tab</p>
    <p>1/ Click on &ldquo;Forwarding and POP/IMAP&rdquo;</p>
    <p></p>
    <p>2/ Click on &ldquo;Add a forwarding address&rdquo;</p>
    <p><img src="images/image8.png" style="max-width:100%"></p>
    <p></p>
    <p></p>
    <p><b>Step 3:</b> Enter your child&rsquo;s personal email address</p>
    <p></p>
    <p>1/ Enter your email and click &ldquo;Next&rdquo;</p>
    <p><img src="images/image12.png" style="max-width:100%"></p>
    <p>2/ Click on &ldquo;Proceed&rdquo; on the prompt</p>
    <p><img src="images/image2.png" style="max-width:100%"></p>
    <p></p>
    <p></p>
    <p><b>Step 4:</b> Enter your confirmation code</p>
    <p>Check your email on your personal address, and enter the code, or click on the link
            in the email :</p>
    <p><img src="images/image10.png" style="max-width:100%"></p>
    <p><img src="images/image5.png" style="max-width:100%"></p>
    <p></p>
    <p><b>Step 5:</b> Enable the forwarding, and save your changes!</p>
    <p>1/ Select &ldquo;Forward &hellip;.&rdquo;</p>
    <p></p>
    <p>2/ Click on &ldquo;Save changes&rdquo;  !</p>
    <p><img src="images/image11.png" style="max-width:100%"></p>



    <h2 id="lostpassword">Retrieving your lost password</h2>
    <p>If you have lost your password, you can <a href="https://support.google.com/accounts/answer/7682439" target="_blank">reset
       it through the password recovery procedure</a>, using the parent&rsquo;s email
       address. We invite you to link an emergency telephone number for more security.</p>
    <p></p>
    <p>NB: If you don&rsquo;t remember your email address, you can get it back using step 1
            at the top of this document (the original password will not work anymore however, once it has been changed,
            so you still need to get a new password using the password recovery procedure).</p>
    <p></p>
    <p></p>

      </div>
    </div>
  </div>





  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/js/bootstrap.bundle.min.js" integrity="sha384-ka7Sk0Gln4gmtz2MlQnikT1wXgYsOg+OMhuP+IlRH9sENBO0LRn5q+8nbTov4+1p" crossorigin="anonymous"></script>

</body>

</html>
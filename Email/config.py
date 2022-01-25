import email.utils
import smtplib
import os, mimetypes
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication


def send_email(to_address, from_address,subject, message,cc=[],files=[],host="localhost", port=587, username='', password=''):
    """Sending an email with python.

       Args:
           to_address (str): the recipient of the email
           from_address (str): the originator of the email
           subject (str): message subject
           message (str): message body
           cc (list[str]): list emails to be copied on the email
           files (list[str]): attachment files
           host (str):the host name
           port (int): the port
           username (str): server auth username
           password (str): server auth password
       """


    today = date.today()
    msg = MIMEMultipart("alternative")
    msg['From'] = email.utils.formataddr(('From Name', from_address))
    msg['To'] = to_address
    msg['Cc'] = ",".join(cc)
    msg['Subject'] = subject+"- {}".format(today.strftime("%b-%Y"))

    to_address = [to_address] + cc

    body = "Good day, \n Please find the attached report for {}".format(
        today.strftime("%b-%Y"))

    html = """\
    <!doctype html>
        <html lang="en">
          <head>
            <title>
            </title>
            <meta http-equiv="X-UA-Compatible" content="IE=edge">
            <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <style type="text/css">
              #outlook a{padding: 0;}
                        .ReadMsgBody{width: 100%;}
                        .ExternalClass{width: 100%;}
                        .ExternalClass *{line-height: 100%;}
                        body{margin: 0; padding: 0; -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%;}
                        table, td{border-collapse: collapse; mso-table-lspace: 0pt; mso-table-rspace: 0pt;}
                        img{border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic;}
                        p{display: block; margin: 13px 0;}
            </style>
            <!--[if !mso]><!-->
            <style type="text/css">
              @media only screen and (max-width:480px) {
                                @-ms-viewport {width: 320px;}
                                @viewport {	width: 320px; }
                            }
            </style>
            <!--<![endif]-->
            <!--[if mso]> 
                <xml> 
                    <o:OfficeDocumentSettings> 
                        <o:AllowPNG/> 
                        <o:PixelsPerInch>96</o:PixelsPerInch> 
                    </o:OfficeDocumentSettings> 
                </xml>
                <![endif]-->
            <!--[if lte mso 11]> 
                <style type="text/css"> 
                    .outlook-group-fix{width:100% !important;}
                </style>
                <![endif]-->
            <style type="text/css">
              @media only screen and (min-width:480px) {
              .dys-column-per-100 {
                width: 100.000000% !important;
                max-width: 100.000000%;
              }
              }
              @media only screen and (min-width:480px) {
              .dys-column-per-100 {
                width: 100.000000% !important;
                max-width: 100.000000%;
              }
              }
            </style>
          </head>
          <body>
            <div>
              <table align='center' border='0' cellpadding='0' cellspacing='0' role='presentation' style='background:#f7f7f7;background-color:#f7f7f7;width:100%;'>
                <tbody>
                  <tr>
                    <td>
                      <div style='margin:0px auto;max-width:600px;'>
                        <table align='center' border='0' cellpadding='0' cellspacing='0' role='presentation' style='width:100%;'>
                          <tbody>
                            <tr>
                              <td style='direction:ltr;font-size:0px;padding:20px 0;text-align:center;vertical-align:top;'>
                                <!--[if mso | IE]>
        <table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td style="vertical-align:top;width:600px;">
        <![endif]-->
                                <div class='dys-column-per-100 outlook-group-fix' style='direction:ltr;display:inline-block;font-size:13px;text-align:left;vertical-align:top;width:100%;'>
                                  <table border='0' cellpadding='0' cellspacing='0' role='presentation' style='vertical-align:top;' width='100%'>
                                    <tr>
                                      <td align='center' style='font-size:0px;padding:10px 25px;word-break:break-word;'>
                                        <div style='color:#4d4d4d;font-family:Oxygen, Helvetica neue, sans-serif;font-size:18px;font-weight:700;line-height:37px;text-align:center;'>
                                          Ronnie The Dev
                                        </div>
                                      </td>
                                    </tr>
                                    <tr>
                                      <td align='center' style='font-size:0px;padding:10px 25px;word-break:break-word;'>
                                        <div style='color:#777777;font-family:Oxygen, Helvetica neue, sans-serif;font-size:14px;line-height:21px;text-align:center;'>
                                          """+message+"""
                                        </div>
                                      </td>
                                    </tr>
                                  </table>
                                </div>
                                <!--[if mso | IE]>
        </td></tr></table>
        <![endif]-->
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </td>
                  </tr>
                </tbody>
              </table>
              <table align='center' border='0' cellpadding='0' cellspacing='0' role='presentation' style='background:#f7f7f7;background-color:#f7f7f7;width:100%;'>
                <tbody>
                  <tr>
                    <td>
                      <div style='margin:0px auto;max-width:600px;'>
                        <table align='center' border='0' cellpadding='0' cellspacing='0' role='presentation' style='width:100%;'>
                          <tbody>
                            <tr>
                              <td style='direction:ltr;font-size:0px;padding:20px 0;text-align:center;vertical-align:top;'>
                                <!--[if mso | IE]>
        <table role="presentation" border="0" cellpadding="0" cellspacing="0"><tr><td style="vertical-align:top;width:600px;">
        <![endif]-->
                                <div class='dys-column-per-100 outlook-group-fix' style='direction:ltr;display:inline-block;font-size:13px;text-align:left;vertical-align:top;width:100%;'>
                                  <table border='0' cellpadding='0' cellspacing='0' role='presentation' style='vertical-align:top;' width='100%'>
                                    <tr>
                                      <td align='center' style='font-size:0px;padding:10px 25px;word-break:break-word;'>
                                        <div style='color:#4d4d4d;font-family:Oxygen, Helvetica neue, sans-serif;font-size:32px;font-weight:700;line-height:37px;text-align:center;'>
                                        </div>
                                      </td>
                                    </tr>
                                    <tr>
                                      <td align='center' style='font-size:0px;padding:10px 25px;word-break:break-word;'>
                                        <div style='color:#777777;font-family:Oxygen, Helvetica neue, sans-serif;font-size:14px;line-height:21px;text-align:center;'>
                                          <strong>
                                           RonnieTheDev|Digital Transformation Developer|
                                            <br />
                                            Company | Information Technology Department
                                            <br />
                                            | +263 777 777 777| +263 444 444 444 |
                                          </strong>
                                          <strong>
                                          </strong>
                                          <br />
                                        </div>
                                      </td>
                                    </tr>
                                  </table>
                                </div>
                                <!--[if mso | IE]>
        </td></tr></table>
        <![endif]-->
                              </td>
                            </tr>
                          </tbody>
                        </table>
                      </div>
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>
          </body>
        </html>
    """

    part1 = MIMEText(body, 'plain')
    part2 = MIMEText(html, 'html')

    # msg.attach(MIMEText(body, 'plain'))

    msg.attach(part1)
    msg.attach(part2)


    for filepath in files:
        ctype, encoding = mimetypes.guess_type(filepath)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        if maintype in ['image', 'audio']:
            add_attachment(filepath)
        else:
            baseName = os.path.basename(filepath)
            att = MIMEApplication(open(filepath, 'rb').read())
            att.add_header('Content-Disposition', 'attachment', filename=baseName)
            msg.attach(att)
            print(filepath, 'added')

    server = smtplib.SMTP(host, port)
    server.ehlo()
    # server.set_debuglevel(3)
    server.starttls()
    server.login(username, password)
    server.sendmail(from_address, to_address, msg.as_string())
    server.quit()
    print("Email has been sent")


def add_attachment(filepath):
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)

    if maintype == 'text':
        fp = open(filepath)
        attachment = MIMEText(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == 'image':
        fp = open(filepath, 'rb')
        attachment = MIMEImage(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == 'audio':
        fp = open(filepath, 'rb')
        attachment = MIMEAudio(fp.read(), _subtype=subtype)
        fp.close()
    else:
        fp = open(filepath, 'rb')
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(attachment)

    baseName = os.path.basename(filepath)
    attachment.add_header('Content-Disposition', 'attachment', filepath=filepath, filename=baseName)
    msg.attach(attachment)
    print(filepath, 'added')
    return 'Done'

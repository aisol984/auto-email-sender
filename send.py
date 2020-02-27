from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.image     import MIMEImage
import smtplib
import xlrd
import re

def is_valid_email(email):
    '''
    This is that check if email address is valid
    '''
    if len(email) > 7:
        return bool(re.match(
            "^.+@(\[?)[a-zA-Z0-9-.]+.([a-zA-Z]{2,3}|[0-9]{1,3})(]?)$", email))

def load_data(file_name):
    '''
    This is that load informations in excel file to send mails
    '''
    sent_members = []
    try:
      wb = xlrd.open_workbook(file_name)
      sheet = wb.sheet_by_index(0)
      for row in range(sheet.nrows):
          if is_valid_email(sheet.row_values(row)[1]):
              sent_members.append(send_mail(sheet.row_values(row)))
    except Exception as e:
      return None
    return sent_members
            

def send_mail(receiver_info):
    '''
    This is that send mail to people in excel file
    '''
    try:
      strFrom = 'info@voteourvoice.com'
      strTo = receiver_info[1]

      msgRoot = MIMEMultipart('related')
      # Mail Subject
      msgRoot['Subject'] = 'VoteOurVoice'
      # Sender address
      msgRoot['From'] = strFrom
      # Receiver address
      msgRoot['To'] = strTo
      msgRoot.preamble = 'This is a multi-part message in MIME format.'

      msgAlternative = MIMEMultipart('alternative')
      msgRoot.attach(msgAlternative)

      msgText = MIMEText('This is the alternative plain text message.')
      msgAlternative.attach(msgText)

      # Email Html Code
      content = """
      <html xmlns="http://www.w3.org/1999/xhtml">
        <head>
          <meta name="viewport" content="width=device-width" />
          <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        </head>
        <body>
          <table style="width: 80%;">
            <tr>
              <td></td>
              <td style="display:block!important;max-width:600px!important;margin:0 auto!important; clear:both!important;">
                <div style="padding:15px; max-width:600px; margin:0 auto; display:block; ">
                <table>
                  <tr>
                    <td>
                      <div>
                        <p>{first_name},</p>
                        <p>My name is Richard Hill and I’m the Founder of VoteOurVoice. My company is creating a Democracy Index which we feel will be a Leading Economic Indicator. The Democracy Index in short will provide data outlining how well people feel they are being represented. The data will provide insight into the likelihood of a regime change (change in ruling political parties). This information will be useful to traders globally and provide clarity for hedging during election cycles.</p>
                        <p>We envision that our Democracy Index will work similar to the University of Michigan Consumer Sentiment Index. The major difference is that our data is not based on polling because our users vote on summaries of the same legislation their elected officials voted on. </p>
                        <p><b>VoteOurVoice Democracy Index</b> – a monthly metric that describes the effectiveness of political representation in the US. It’s the weighted average of performance metrics (tallied by VoteOurVoice data for each of 50 states) for all Members of the House and Senate.</p>
                        <p>I am reaching out to you to invite you to register with VoteOurVoice and in order to experience the benefits of providing you politician performance data. If you would like to discuss our platform, please do not hesitate to reach out. If you register with the next 24 hours, you will receive 25% OFF your first data report.</p>
                        </div>
                      <table style="background-color: black;" width="100%">
                        <tr>
                          <td>									
                            <table style="width: 100%;">
                              <tr>
                                <td colspan="3">
                                  <h1 style="text-align: center; color: #fff; width: 100%;font-family: cursive;">Register Today:Offer Valid for first report!</h1>
                                </td>
                              </tr>
                              <tr>
                                <td colspan="3"></td>
                              </tr>
                              <tr>
                                <td style="background-color: white; width: 40%;">
                                  <a href="https://www.voteourvoice.com" style="text-decoration: none;">
                                    <span style="color: black; display: block; width: 100%; text-align: center;font-family: Arial, Helvetica, sans-serif;font-size: large"><b>Save 25% Today!</b></span>
                                  </a>
                                </td>
                                <td style="width: 2%;"></td>
                                <td style="background-color: white; width: 40%;">
                                  <a href="https://www.voteourvoice.com" style="text-decoration: none;">
                                    <span style="color: black; display: block; width: 100%; text-align: center;font-family:  Arial, Helvetica, sans-serif;font-size: large"><b>Democracy Index Report</b></span>
                                  </a>
                                </td>
                              </tr>
                              <tr>
                                <td colspan="3"></td>
                              </tr>
                            </table>
                          </td>
                        </tr>
                      </table>
                      <p>Here is the information about my startup and me:</p>
                      <p><b>About VoteOurVoice</b><br>VoteOurVoice is a global, nonpartisan, voting tool that gives people greater transparency in their democracy by allowing them to grade the performance of their politicians. VoteOurVoice is a Public Benefit Corporation and our mission is to provide voters with an easy and objective way to make informed decisions about their electoral choices and hold their politicians accountable for their performance while in office. We are solving a problem that most voters have...insufficient data to evaluate politicians. Users act like elected officials by ‘voting’ on summaries of actual legislation that their particular representative voted on during the preceding four years. In this interactive web-based application, the citizen’s votes are compared to those recorded by their representatives, and the report card is generated. The report card shows how well the politician represented the individual user and also the entire voting district.</p>
                      <p><b>Richard Hill | Bio</b><br>Richard is a social entrepreneur who is Founder & CEO of VoteOurVoice, and Chairman of Hill Industries, an investment company that makes principal investments in privately held companies in manufacturing, construction, and technology industries. He has 20 years combined fixed income trading, investment banking ad consulting, real estate, and business development experience in corporate finance and capital markets. Mr. Hill worked at Citigroup (as a VP in Structured Finance), Barclays Capital (as a bond trader), and CBRE (as an industrial real estate broker). He completed structured finance transactions valued at $16.5 billion, had management responsibilities for trading books valued at over $300 million, and marketed 1.4 million square feet of commercial real estate. He was also a senior partner in the merchant banking firm, Quantum Capital Partners, LLC, where he was the Financial Advisor to Forest City Ratner Companies for their $4.5 billion mixed use project (which includes the 20,000 seat Barclays Center) in Brooklyn, New York.  Mr. Hill has a Master’s degree in Applied Economics from New York University, an MBA from California State Polytechnic University, Pomona, and a BA in Economics from the University of California, Irvine.</p>
                      <p>Thank you,</p>
                      <div>Richard Hill<br><b>CEO | VoteOurVoice</b><br><a href="https://www.voteourvoice.com">www.voteourvoice.com</a><p></p></div>
                      <img src="cid:image1">
                    </td>
                  </tr>
                </div>
              </td>
            </tr>
          </table>
        </body>
        </html>
      """.format(first_name=receiver_info[0]) # Add first Name
      msgText = MIMEText(content, 'html')
      msgAlternative.attach(msgText)

      # Add company logo on email html
      fp = open('logo.png', 'rb')
      msgImage = MIMEImage(fp.read())
      fp.close()

      msgImage.add_header('Content-ID', '<image1>')
      msgRoot.attach(msgImage)

      # Send the email (this example assumes SMTP authentication is required)
      smtp = smtplib.SMTP('smtp.gmail.com', 587)
      smtp.ehlo()
      smtp.starttls()
      smtp.login(strFrom, 'morgan1401')
      smtp.sendmail(strFrom, strTo, msgRoot.as_string())
      smtp.quit()
    except Exception as e:
      return None
    return receiver_info[0]
import smtplib
import os.path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pyodbc
import xlwt

def send_email(subject, message, from_email, to_email=[], attachment=[]):
    """
    :param subject: email subject
    :param message: Body content of the email (string), can be HTML/CSS or plain text
    :param from_email: Email address from where the email is sent
    :param to_email: List of email recipients, example: ["a@a.com", "b@b.com"]
    :param attachment: List of attachments, exmaple: ["file1.txt", "file2.txt"]
    """
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = from_email
    msg['To'] = ", ".join(to_email)
    msg.attach(MIMEText(message, 'html'))

    for f in attachment:
        with open(f, 'rb') as a_file:
            basename = os.path.basename(f)
            part = MIMEApplication(a_file.read(), Name=basename)

        part['Content-Disposition'] = 'attachment; filename="%s"' % basename
        msg.attach(part)

    email = smtplib.SMTP('aphubtran01.ap-thai.com',25)
    email.sendmail(from_email, to_email, msg.as_string())
    email.quit()
    return;

def GenData2Xls():
    book = xlwt.Workbook()
    sheet1 = book.add_sheet("PySheet1")
    
    # Grey background for the header row
    BkgPat = xlwt.Pattern()
    BkgPat.pattern = xlwt.Pattern.SOLID_PATTERN
    BkgPat.pattern_fore_colour = 22
    
    # Bold Fonts for the header row
    font = xlwt.Font()
    font.name = 'Calibri'
    font.bold = True
    
    # Non-Bold fonts for the body
    font0 = xlwt.Font()
    font0.name = 'Calibri'
    font0.bold = False
    
    # style and write field labels
    style = xlwt.XFStyle()
    style.font = font
    style.pattern = BkgPat
    
    style0 = xlwt.XFStyle()
    style0.font = font0
    
    connection = pyodbc.connect('Driver={SQL Server};Server=192.168.2.58;Database=db_iconcrm_fusion;uid=iconuser;pwd=P@ssw0rd')   
    cursor = connection.cursor()  
     
    strSQL = "SELECT \
      a.ProductID \
      ,a.Project \
      ,isnull(a.SAPProductID,'-') as SAPProductID \
      ,isnull(a.SAPCostCenter,'-') as SAPCostCenter \
      ,b.CompanyNameThai \
      ,case ISNULL(RTPExcusive,1) \
        when 1 then 'Active' \
        when 2 then 'Active' \
        else 'Inactive' \
    end  as StatusProject \
    FROM dbo.ICON_EntForms_Products a left join \
    dbo.ICON_EntForms_Company b ON a.CompanyID = b.CompanyID \
    WHERE 1=1 \
    ORDER BY ProductID "
    #Modified by Suchat S. 2018-07-02 change ISNULL(RTPExcusive,1)
    #AND a.RTPExcusive IS NOT NULL" 
  
    cursor.execute(strSQL)
    result_set = cursor.fetchall()
    
    cols = ["ProductID","Project","SAPProductID","SAPCostCenter","CompanyNameThai","StatusProject"]   
    #Gen Header Data
    for colnum, value in enumerate(cols):
        sheet1.write(0, colnum, value,style)
    
    # Genearate Detail Data
    row_number=1
    for row in result_set:
        column_num=0
        for item in row:
            #sheet1.write(row_number,column_num,str(item))
            sheet1.write(row_number,column_num,str(item),style0)
            column_num=column_num+1
    
        row_number=row_number+1
    
    book.save("projectSAP.xls")
    
    cursor.close()
    del cursor
    
    print("Generate test.xls File Successful..!!")

print("<<< Generate Data to Excel File Start >>>")
GenData2Xls()
print("<<< Generate Data to Excel File Finish >>>")

print("<<< Send Mail Start>>>")
sender = 'SysMail@apthai.com'
receivers = ['project_code@apthai.com','sutthikarn_c@apthai.com','tanonchai@apthai.com','suchat_s@apthai.com']
#receivers = ['tanonchai@apthai.com','suchat_s@apthai.com']
#receivers = ['suchat_s@apthai.com']

#subject = "Send mail Test SMTP From [" + socket.gethostname() + "]"
subject = "[CRM] Project Code Mapping SAP"
message = """\
<html>
  <head></head>
  <body>
    <p style="font-family:AP;">Dear All<br>
       &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This e-mail auto send from system and attached excel file for project code mapping with SAP code.</p>
    <p style="font-family:AP;">Best Regards,<br>
    IT Team.</p>
  </body>
</html>
"""
attachedFile = ['projectSAP.xls']

send_email(subject, message, sender, receivers, attachedFile)
print("Successfully sent email")
print("<<< Send Mail Finish>>>")
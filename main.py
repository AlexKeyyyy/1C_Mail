import smtplib
import time
from email.header import Header
from email.mime.text import MIMEText
from datetime import timedelta
from bson.objectid import ObjectId
from pymongo import MongoClient
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from reportlab.pdfgen import canvas
import os
from pymongo import MongoClient
from bson import ObjectId


def make_task_report(user, task, user_task):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    styles = getSampleStyleSheet()

    title_image_path = './uploads/PDF/ReportTitle.png'
    background_image_path = './uploads/PDF/ReportBack.png'
    code_image_path = './uploads/PDF/ReportCode.png'

    title_image_path = title_image_path if os.path.exists(title_image_path) else None
    background_image_path = background_image_path if os.path.exists(background_image_path) else None
    code_image_path = code_image_path if os.path.exists(code_image_path) else None

    elements = []
    if title_image_path:
        elements.append(Image(title_image_path, width=doc.width, height=doc.height))

    title = Paragraph(f"Отчет по задаче №{task['taskNumber']}", styles['Title'])
    elements.append(title)
    elements.append(Spacer(1, 12))

    user_info = f"Кандидат: {user['surname']} {user['name']} {user.get('patro', '')}<br/>Email: {user['email']}"
    user_paragraph = Paragraph(user_info, styles['Normal'])
    elements.append(user_paragraph)
    elements.append(Spacer(1, 12))

    task_errors = 0
    task_vulnerabilities = 0
    task_defects = 0

    for issue in user_task['results']['issues']:
        if issue['tags'] and issue['message'] != "Нужно заменить символ неразрывного пробела на обычный пробел":
            for tag in issue['tags']:
                if tag == "error":
                    task_errors += 1
                if tag == "badpractice":
                    task_defects += 1

    total_issues = task_errors + task_defects + task_vulnerabilities
    stats = f"Всего ошибок: {total_issues}"
    stats_paragraph = Paragraph(stats, styles['Normal'])
    elements.append(stats_paragraph)
    elements.append(Spacer(1, 12))

    subject = f"[ОЭ {user['_id']}][{task['taskNumber']}][SonarQube] Запрос поддержки от Admin"
    body = f"Описание проблемы: (опишите, что случилось)\nПриоритет: от 1 до 4 (1 - срочный, 4 - некритичный)\nЖелаемая дата окончания сопровождения: (проставьте желаемую дату разрешения вопроса)\n\nС уважением,\n{user['surname']} {user['name']} {user.get('patro', '')}"
    mailto_link = f"mailto:VCP@mail.ru?subject={subject}&body={body}"
    support_paragraph = Paragraph(f'<a href="{mailto_link}">ПОДДЕРЖКА</a>', styles['Normal'])
    elements.append(support_paragraph)
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Ошибки", styles['Heading2']))
    if not user_task['results']['issues']:
        elements.append(Paragraph("Ошибки отсутствуют.", styles['Normal']))
    else:
        for i, issue in enumerate(user_task['results']['issues'], start=1):
            if issue['message'] != "Нужно заменить символ неразрывного пробела на обычный пробел":
                severity_color = {'INFO': 'green', 'MINOR': 'orange', 'CRITICAL': 'red'}.get(issue['severity'], 'black')
                issue_text = f'<font color="{severity_color}">Ошибка {i} [{issue["severity"]}]</font><br/>Описание: {issue["message"]}<br/>Строка: {issue["line"]}'
                elements.append(Paragraph(issue_text, styles['Normal']))
                elements.append(Spacer(1, 6))

    elements.append(Paragraph("Исходный код", styles['Heading2']))
    if user_task['codeText']:
        code_lines = user_task['codeText'].split("\n")
        for line_number, line in enumerate(code_lines, start=1):
            line_color = 'black'
            for issue in user_task['results']['issues']:
                if issue['line'] == line_number:
                    line_color = {'INFO': 'green', 'MINOR': 'orange', 'CRITICAL': 'red'}.get(issue['severity'], 'black')
                    break
            elements.append(
                Paragraph(f'{str(line_number).zfill(4)}: <font color="{line_color}">{line}</font>', styles['Code']))
            elements.append(Spacer(1, 6))
    else:
        elements.append(Paragraph("Код отсутствует.", styles['Normal']))

    doc.build(elements)
    buffer.seek(0)
    return buffer


def save_and_send_report(user, task_document, user_task, admin_email):
    pdf_buffer = make_task_report(user, task_document, user_task)
    report_filename = f"user_task_report_{user['_id']}_{task_document['taskNumber']}.pdf"

    with open(report_filename, 'wb') as f:
        f.write(pdf_buffer.read())

    subject = "Отчет по задаче"
    message = "Вложение содержит отчет по задаче."

    msg = MIMEMultipart()
    msg['From'] = "alexkey2017@yandex.ru"
    msg['To'] = admin_email
    msg['Subject'] = subject
    msg.attach(MIMEText(message, 'plain'))

    with open(report_filename, 'rb') as attachment:
        part = MIMEApplication(attachment.read(), Name=report_filename)
        part['Content-Disposition'] = f'attachment; filename="{report_filename}"'
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login("alexkey2017@yandex.ru", "daqsyf-fEdxo7-fydfuh")
        server.sendmail("alexkey2017@yandex.ru", admin_email, msg.as_string())
        server.close()
        print(f"Отчет отправлен на {admin_email}")
    except Exception as e:
        print(f"Ошибка отправки: {e}")


def send_email(recepient_email, subject, msg_text, recepient_name):
    login = 'alexkey2017@yandex.ru'
    password = 'daqsyf-fEdxo7-fydfuh'

    # Укажите путь или URL к логотипу
    logo_url = 'https://example.com/logo.png'

    html_msg = f"""
        <!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
        <html>
        <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
            <title>Нотификация</title>
            <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,600&display=swap"
                em-class="em-font-Montserrat-SemiBold">
            <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,400&display=swap"
                em-class="em-font-Montserrat-Regular">
            <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,800&display=swap"
                em-class="em-font-Montserrat-ExtraBold">
            <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Montserrat:ital,wght@0,700&display=swap"
                em-class="em-font-Montserrat-Bold">
            <style type="text/css">
                html {{
                    -webkit-text-size-adjust: none;
                    -ms-text-size-adjust: none;
                }}
            </style>
            <style em="styles">
                .em-font-Montserrat-Bold {{
                    font-family: Montserrat, sans-serif !important;
                    font-weight: 700 !important;
                }}
                .em-font-Montserrat-ExtraBold {{
                    font-weight: 800 !important;
                }}
                .em-font-Montserrat-ExtraBold, .em-font-Montserrat-Medium, .em-font-Montserrat-Regular {{
                    font-family: Montserrat, sans-serif !important;
                }}
                .em-font-Montserrat-Regular {{
                    font-weight: 400 !important;
                }}
                .em-font-Montserrat-SemiBold {{
                    font-family: Montserrat, sans-serif !important;
                    font-weight: 600 !important;
                }}
                @media only screen and (max-device-width:660px), only screen and (max-width:660px) {{
                    .em-narrow-table {{
                        width: 100% !important;
                        max-width: 660px !important;
                        min-width: 280px !important;
                    }}
                    .em-mob-line_height-25px {{
                        line-height: 25px !important;
                    }}
                    .em-mob-wrap {{
                        display: block !important;
                    }}
                    .em-mob-width-100perc {{
                        width: 100% !important;
                        max-width: 100% !important;
                        min-width: 100% !important;
                    }}
                    .em-mob-padding_top-20 {{
                        padding-top: 20px !important;
                    }}
                    .em-mob-padding_bottom-20 {{
                        padding-bottom: 20px !important;
                    }}
                    .em-mob-padding_bottom-14 {{
                        padding-bottom: 14px !important;
                    }}
                    .em-mob-font_size-17px {{
                        font-size: 17px !important;
                    }}
                    .em-mob-font_size-14px {{
                        font-size: 14px !important;
                    }}
                    .em-mob-padding_left-20 {{
                        padding-left: 20px !important;
                    }}
                    .em-mob-padding_right-0 {{
                        padding-right: 0 !important;
                    }}
                    .em-mob-padding_left-0 {{
                        padding-left: 0 !important;
                    }}
                    .em-mob-vertical_align-middle {{
                        vertical-align: middle !important;
                    }}
                    .em-mob-height-auto {{
                        height: auto !important;
                    }}
                    .em-mob-height-48px {{
                        height: 48px !important;
                    }}
                    .em-show-td-desktop {{
                        display: none !important;
                    }}
                    .em-mob-padding_bottom-44 {{
                        padding-bottom: 44px !important;
                    }}
                    .em-mob-padding_top-45 {{
                        padding-top: 45px !important;
                    }}
                    .em-mob-font_size-12px {{
                        font-size: 12px !important;
                    }}
                    .em-mob-padding_right-20 {{
                        padding-right: 20px !important;
                    }}
                    .em-mob-width-342px {{
                        width: 342px !important;
                        max-width: 342px !important;
                        min-width: 342px !important;
                    }}
                    .em-mob-width-345px {{
                        width: 345px !important;
                        max-width: 345px !important;
                        min-width: 345px !important;
                    }}
                    .em-mob-text_align-center {{
                        text-align: center !important;
                    }}
                    .em-mob-width-336px {{
                        width: 336px !important;
                        max-width: 336px !important;
                        min-width: 336px !important;
                    }}
                }}
            </style>
        </head>

        <body style="margin: 0; padding: 0;">
            <span class="preheader"
                style="display: none !important; visibility: hidden; opacity: 0; color: #F8F8F8; height: 0; width: 0; font-size: 1px;">&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;‌&nbsp;</span>
            <div style="font-size:0px;color:transparent;opacity:0;">
        ⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀⠀
            </div>
            <table cellpadding="0" cellspacing="0" border="0" width="100%" style="font-size: 1px; line-height: 28px;">
                <tr em="group">
            <td align="center"
                background="https://img.us5-usndr.com/en/v5/user-files?userId=6945886&resource=himg&disposition=inline&name=6dafp9o6qug5bpt439aew759pxrzbjyn1jferau1grnh6yiif831jzhg4jyqekerunfr6zz38n8pj8o8bf96n9ja8c8ufwsthpjej7nqmpf8ofpq8hokt68gtwaetamyzzsjgftcyy1hp74z41pd4eb4qte"
                style="background-size: cover; padding: 40px 0px 48px; background-color: #e22d2d;"
                class="em-mob-padding_top-20 em-mob-padding_bottom-20 em-mob-padding_left-20 em-mob-padding_right-20"
                bgcolor="#E22D2D">
                <table cellpadding="0" cellspacing="0" width="100%" border="0"
                    style="max-width: 660px; min-width: 660px; width: 660px;" class="em-narrow-table">
                    <tr em="block" class="em-structure">
                        <td align="left" valign="middle"
                            class="em-mob-height-auto em-mob-padding_top-20 em-mob-padding_right-20 em-mob-padding_bottom-20 em-mob-padding_left-20"
                            style="background-position: 0% center; padding: 40px 28px 40px 40px; background-repeat: no-repeat; background-size: cover; border-top-left-radius: 16px; border-top-right-radius: 16px;"
                            background="https://img.us5-usndr.com/en/v5/user-files?userId=6945886&resource=himg&disposition=inline&name=6jjz6n57fnfurzt439aew759pxrzbjyn1jferau1grnh6yiif831jzhg4jyqekeruzu1sqd4k9r6ffn533o778eh3ejruhqq6fob1ko3af8w7f5efqwyqgqp1t1chpo4e836rdemxf9tad9cr8jrwnjnjic">
                            <table border="0" cellspacing="0" cellpadding="0" class="em-mob-width-100perc">
                                <tr>
                                    <td width="592" class="em-mob-wrap em-mob-width-100perc">




                                        <table cellpadding="0" cellspacing="0" border="0" width="46%" em="atom"
                                            class="em-mob-width-345px">
                                            <tr>
                                                <td style="padding-right: 0px; padding-left: 0px;"
                                                    class="em-mob-vertical_align-middle">
                                                    <div style="font-family: Helvetica, Arial, sans-serif; font-size: 24px; line-height: 32px; color: #333333;"
                                                        class="em-font-Montserrat-Bold em-mob-text_align-center"><img
                                                            src="https://img.us5-usndr.com/en/v5/user-files?userId=6945886&resource=himg&disposition=inline&name=6xgecqto1xyjbit439aew759pxrzbjyn1jferau1grnh6yiif831jzhg4jyqekeruqx4i8uxkdn5gqm3b3wrs1ux9b6p9zwifrnzdnr73xg1oxfxfpnkd5p3wg5d48jufkykgr3trtkqmi4z41pd4eb4qte"
                                                            width="30" border="0" alt=""
                                                            style="display: inline-block; vertical-align: text-bottom; width: 100%; max-width: 30px;">&nbsp;<strong>Technologies&nbsp;</strong>
                                                    </div>
                                                </td>
                                            </tr>
                                        </table>
                                        <table cellpadding="0" cellspacing="0" border="0" width="100%" em="atom"
                                            class="em-mob-width-336px">
                                            <tr>
                                                <td style="padding: 1px 0px 0px;">
                                                    <div style="font-family: Helvetica, Arial, sans-serif; font-size: 10px; line-height: 17px; color: #898787"
                                                        class="em-font-Montserrat-Regular em-mob-text_align-center">
                                                        Платформа проверки тестовых заданий&nbsp;</div>
                                                </td>
                                            </tr>
                                        </table>
                                        <table cellpadding="0" cellspacing="0" border="0" width="100%" em="atom">
                                            <tr>
                                                <td style="padding: 1px 0px 0px;" class="em-show-td-desktop">
                                                    <div
                                                        style="font-family: -apple-system, 'Segoe UI', 'Helvetica Neue', Helvetica, Roboto, Arial, sans-serif; font-size: 20px; line-height: 17px; color: #333333;">
                                                        &nbsp;</div>
                                                </td>
                                            </tr>
                                        </table>

                                        <table cellpadding="0" cellspacing="0" border="0" width="100%" em="atom"
                                            class="em-mob-width-342px">
                                            <tr>
                                                <td style="padding: 17px 0px 5px;"
                                                    class="em-mob-text_align-center em-mob-padding_bottom-44 em-mob-padding_top-45">
                                                    <div style="font-family: Helvetica, Arial, sans-serif; font-size: 38px; line-height: 42px; color: #E22D2D"
                                                        class="em-font-Montserrat-ExtraBold">Уведомление&nbsp;</div>
                                                </td>
                                            </tr>
                                        </table>
                                        <table cellpadding="0" cellspacing="0" border="0" em="atom" width="55%"
                                            class="em-mob-width-342px">
                                            <tr>
                                                <td style="padding-bottom: 5px; padding-top: 10px;"
                                                    class="em-mob-text_align-center em-mob-padding_bottom-14">
                                                    <div style="font-family: Helvetica, Arial, sans-serif; font-size: 14px; line-height: 20px; color: #000000;"
                                                        class="em-mob-line_height-25px em-font-Montserrat-SemiBold em-mob-font_size-17px">
                                                        Приветствуем, <span style="color: #eb0f0f;">{recepient_name}</span>!<br>
                                                    </div>
                                                </td>
                                            </tr>
                                        </table>
                                        <table cellpadding="0" cellspacing="0" border="0" em="atom" width="55%"
                                            class="em-mob-width-345px">
                                            <tr>
                                                <td style="padding-bottom: 26px;" class="em-mob-text_align-center">
                                                    <div style="font-family: Helvetica, Arial, sans-serif; font-size: 10px; line-height: 20px; color: #000000"
                                                        class="em-mob-line_height-25px em-font-Montserrat-Regular em-mob-font_size-12px em-mob-text_align-center">
                                                        {msg_text}<br></div>
                                                </td>
                                            </tr>
                                        </table>
                                        <table cellpadding="0" cellspacing="0" border="0" em="atom" width="60%"
                                            class="em-mob-width-100perc">
                                            <tr>
                                                <td style="padding-bottom: 26px;">
                                                    <div style="font-family: Helvetica, Arial, sans-serif; font-size: 12px; line-height: 20px; color: #000000"
                                                        class="em-mob-line_height-25px em-font-Montserrat-SemiBold em-mob-text_align-center em-mob-font_size-14px">
                                                        С уважением,<br>команда BIA Technologies<br></div>
                                                </td>
                                            </tr>
                                        </table>

                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>






                    <tr em="block" class="em-structure">
                        <td align="center"
                            style="padding-top: 25px; padding-right: 30px; padding-left: 30px; background-color: #000000; background-repeat: repeat; border-bottom-right-radius: 16px; border-bottom-left-radius: 16px;"
                            class="em-mob-padding_top-20 em-mob-padding_right-20 em-mob-padding_bottom-20 em-mob-padding_left-20"
                            bgcolor="#000000">
                            <table border="0" cellspacing="0" cellpadding="0" class="em-mob-width-100perc">
                                <tr>


                                    <td width="600" valign="top" class="em-mob-wrap em-mob-width-100perc">
                                        <table cellpadding="0" cellspacing="0" border="0" width="100%" em="atom">
                                            <tr>
                                                <td align="center"
                                                    class="em-mob-padding_right-0 em-mob-padding_left-0 em-mob-vertical_align-middle em-mob-height-48px">
                                                    <table border="0" cellspacing="0" cellpadding="0">
                                                        <tr>
                                                            <td width="40" align="center">
                                                                <a href="https://vk.com/biatech" target="_blank"><img
                                                                        src="https://img.us5-usndr.com/en/v5/user-files?userId=6945886&resource=himg&disposition=inline&name=6ffjc833ioishzt439aew759pxrzbjyn1jferau1grnh6yiif831jzhg4jyqekerurummqpbce78fazby3sfgkzrgj5jntw3djdw9owdd3k71duscd58hhzz4u18gijpcomethfxgbxiizsy6nux6yrzoiw"
                                                                        width="23" border="0" alt=""
                                                                        style="display: block; max-width: 23px;"></a>
                                                            </td>
                                                            <td width="40" align="center">
                                                                <a href="https://t.me/biatechnologies"
                                                                    target="_blank"><img
                                                                        src="https://img.us5-usndr.com/en/v5/user-files?userId=6945886&resource=himg&disposition=inline&name=68jnh1ghd1xsmmt439aew759pxrzbjyn1jferau1grnh6yiif831jzhg4jyqekeruiygkm11tzqa6uj4ymy7arfoyuot8utez7crka5pc7mct8buw8khbhwaakx35swsz334ch8pk3csn5z7chwe3wx7ekh"
                                                                        width="22" border="0" alt=""
                                                                        style="display: block; max-width: 22px; width: 100%;"></a>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                            <table border="0" cellspacing="0" cellpadding="0" class="em-mob-width-100perc">
                                <tr>
                                    <td width="600" valign="middle" class="em-mob-wrap em-mob-width-100perc"
                                        align="center" style="border-width: 1px; border-color: #e5e5e5;">
                                        <table cellpadding="0" cellspacing="0" border="0" width="70%" em="atom"
                                            class="em-mob-width-100perc">
                                            <tr>
                                                <td
                                                    style="padding-top: 20px; padding-bottom: 16px; border-top-width: 1px; border-top-color: #eeeeee;">
                                                    <div style="font-family: Helvetica, Arial, sans-serif; font-size: 9px; line-height: 21px; color: #807e7e;"
                                                        align="center"
                                                        class="em-font-Montserrat-Regular em-mob-text_align-center em-mob-font_size-12px">
                                                        <span style="color: #d5d4d4;">Это автоматически созданное
                                                            письмо. Не нужно отвечать на него</span>.<br><br><a
                                                            href="mailto:vcp-tech-sup@mail.ru?subject=[ОЭ][Email]"
                                                            target="_blank"
                                                            style="color: #1C52DC; text-decoration:  underline;"><span
                                                                style="color: #ffffff; text-decoration: underline;">Служба
                                                                поддержки</span></a></div>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        </table>
        </body>
        <div class="XTranslate"></div>
        </html>
"""

    msg = MIMEText(html_msg, 'html', 'utf-8')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = login
    msg['To'] = recepient_email

    s = smtplib.SMTP('smtp.yandex.ru', 587, timeout=10)

    try:
        s.starttls()
        s.login(login, password)
        s.send_message(msg)

        print(f"Email successfully sent to {recepient_email}")
        return True
    except Exception as ex:
        print(f"Failed to send email to {recepient_email}. Error: {ex}")
        return False
    finally:
        s.quit()


# Настройки MongoDB
mongo_client = MongoClient(
    'mongodb+srv://alexykoba:u0PQW4fxB2mHUtb9@cluster0.rpznl0k.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0')
db = mongo_client['test']
collection = db['usertasks']
collectionTasks = db['tasks']
admins_collection = db['admins']
users_collection = db['users']
tasks_collection = db['tasks']

def check_and_send_email():
    while True:
        try:
            # Scenario 1: Email to User when mark is not -1 and status is "graded"
            documents = collection.find({"mark": {"$ne": -1}, "mark_email": {"$ne": "sent"}, "sonarStatus": "checked", "status": "graded"})
            for document in documents:
                user_id = document['user_id']
                task_id = document['task_id']
                mark = document['mark']
                document_id = document['_id']
                code_text = document.get('codeText', 'Код отсутствует')

                # Получение email пользователя по user_id
                user_document = db['users'].find_one({"_id": ObjectId(user_id)})
                task_document = db['tasks'].find_one({"_id": ObjectId(task_id)})
                if not user_document or not task_document:
                    continue

                user_email = user_document['email']
                user_name = f"{user_document['surname']} {user_document['name']} {user_document['patro']}"
                task_number = task_document['taskNumber']
                done_at = document['doneAt'] + timedelta(hours=3)

                formatted_done_at = done_at.strftime('%Y-%m-%d %H:%M:%S')

                # Отправка письма пользователю
                message_text = (
                                f"Ваше решение задания {task_number}, отправленное {formatted_done_at} (по МСК) было оценено экспертом. "
                                f"Вы можете ознакомиться с его заключением в своем профиле личного кабинета.")

                if send_email(user_email, 'Ваша оценка была изменена', message_text, user_name):
                    # Обновление документа, чтобы предотвратить повторные отправки
                    collection.update_one({"_id": ObjectId(document_id)}, {"$set": {"mark_email": "sent"}})
                    print("Оценка была изменена (юзеру)")

            # Scenario 2: Email to Admin when mark is -1 and status is "checking"
            admin_documents = collection.find({"mark": -1, "sonar_email": {"$ne": "sent"}, "sonarStatus": "checked", "status": "checking"})
            for document in admin_documents:
                user_id = document['user_id']
                user_document = db['users'].find_one({"_id": ObjectId(user_id)})
                user_email = user_document['email']
                task_id = document['task_id']
                document_id = document['_id']
                code_text = document.get('codeText', 'Код отсутствует')

                # Получение информации о пользователе и задании
                user_document = db['users'].find_one({"_id": ObjectId(user_id)})
                task_document = db['tasks'].find_one({"_id": ObjectId(task_id)})
                if not user_document or not task_document:
                    continue

                user_name = f"{user_document['surname']} {user_document['name']} {user_document['patro']}"
                task_number = task_document['taskNumber']
                done_at = document['doneAt'] + timedelta(hours=3)

                formatted_done_at = done_at.strftime('%Y-%m-%d %H:%M:%S')


                # Отправка письма админу

                admin_name = 'Admin'
                message_text = (f"Задание {task_number} от {user_name}, отправленное {formatted_done_at} (по МСК), "
                                f"было проверено и готово к экспертной оценке. "
                                f"В личном кабинете Вы также может ознакомиться с отчетом анализа решения и кодом кандидата для проставления экспертного заключения.")

                message_text_to_user = (
                    f"Задание {task_number}, отправленное {formatted_done_at} (по МСК), "
                    f"было проверено и готово к экспертной оценке. "
                    f"Пока эксперт не оставил обратной связи, Вы можете ознакомиться с отчетом анализа решения.")

                send_email(user_email, "Задание готово к проверке", message_text_to_user, user_name)
                print("Задание готово к проверке (юзеру)")
                admin_emails = admins_collection.distinct("email")
                for admin_email in admin_emails:
                    if send_email(admin_email, 'Задание готово к проверке', message_text, admin_name):
                        # Обновление документа, чтобы предотвратить повторные отправки
                        collection.update_one({"_id": ObjectId(document_id)}, {"$set": {"sonar_email": "sent"}})
                        print("Задание готово к проверке (админу)")

            # Scenario 3: Email to user about new tasks
            users = collection.aggregate([
                {"$match": {"status": "assigned", "new_task_email": {"$ne": "sent"}}},
                {"$group": {"_id": "$user_id", "tasks": {"$push": "$task_id"}}}
            ])


            for user in users:
                user_id = user['_id']
                task_ids = user['tasks']

                # Получение данных пользователя
                user_document = db['users'].find_one({"_id": ObjectId(user_id)})
                if not user_document:
                    continue

                user_name = f"{user_document['surname']} {user_document['name']}"
                user_email = user_document['email']

                # Получение номеров заданий
                task_numbers = []
                for task_id in task_ids:
                    task_document = db['tasks'].find_one({"_id": ObjectId(task_id)})
                    if task_document:
                        task_numbers.append(task_document['taskNumber'])

                if task_numbers:
                    task_numbers_sorted = sorted(task_numbers)
                    task_numbers_str = ', '.join(f"{num}" for num in task_numbers_sorted)
                    days_to_complete = 7  # Задайте необходимое количество дней
                    message_text = (
                        f"На Вас назначено новое(ые) задание(ия): \n"
                        f"{task_numbers_str}. \n"
                        f"Вы можете приступить к их выполнению в своем личном кабинете! \n"
                        f"Обратите внимание, что на выполнение тестовых заданий выдается {days_to_complete} дней. \n"
                        f"Удачи!"
                    )

                    if send_email(user_email, 'Назначены новые задания', message_text, user_name):
                        # Обновление документов, чтобы предотвратить повторные отправки
                        collection.update_many(
                            {"user_id": user_id, "status": "assigned", "new_task_email": {"$ne": "sent"}},
                            {"$set": {"new_task_email": "sent"}}
                        )
                        print("Назначены новые задания (юзеру)")

            # Scenario 4: Email to admins about new candidates
            users = db['users'].find({"role": "user", "new_user_email": {"$ne": "sent"}})
            for user in users:
                user_full = f"{user['surname']} {user['name']} {user['patro']}"
                user_email = user['email']
                user_id = user['_id']

                message_text = (
                    f"На платформе проверки тестовых заданий зарегистрирован новый кандидат.\n"
                    f"ФИО: {user_full} \n"
                    f"E-mail: {user_email} \n"
                )
                admin_emails = admins_collection.distinct("email")
                for admin_email in admin_emails:
                    if send_email(admin_email, "Зарегистрирован новый кандидат", message_text, "Admin"):
                        db['users'].update_one({"_id": ObjectId(user_id)}, {"$set": {"new_user_email": "sent"}})
                        print("Зарегистрирован новый кандидат (админу)")

            # Scenario 5: Находим пользователей, у которых есть задания и письмо "all_task_email" еще не было отправлено
            users = collection.distinct("user_id", {"all_task_email": {"$ne": "sent"}})
            for user_id in users:
                user_tasks = list(collection.find({"user_id": user_id}))
                all_checked = all(task['sonarStatus'] == "checked" for task in user_tasks)

                if all_checked and user_tasks:
                    user_document = users_collection.find_one({"_id": ObjectId(user_id)})
                    if not user_document:
                        continue

                    user_name = f"{user_document['surname']} {user_document['name']} {user_document.get('patro', '')}"
                    task_ids = [task['task_id'] for task in user_tasks]
                    task_documents = [tasks_collection.find_one({"_id": ObjectId(task_id)}) for task_id in task_ids]

                    task_numbers_sorted = sorted(
                        (task_doc['taskNumber'] for task_doc in task_documents if task_doc),
                        key=int
                    )

                    message_text = (
                        f"Пользователь {user_name} завершил все выданные Вами тестовые задания. "
                        f"Назначьте новое задание для продолжения тестирования или свяжитесь с кандидатом."
                    )

                    admin_emails = admins_collection.distinct("email")
                    for admin_email in admin_emails:
                        # Генерация и отправка отчета
                        for task_document, user_task_document in zip(task_documents, user_tasks):
                            if task_document and user_task_document:
                                pdf_buffer = make_task_report(user_document, task_document, user_task_document)
                                report_filename = f"user_task_report_{user_document['_id']}_{task_document['taskNumber']}.pdf"

                                with open(report_filename, 'wb') as f:
                                    f.write(pdf_buffer.read())

                                if send_email(admin_email, 'Пользователь завершил задания', message_text, "Admin"):
                                    collection.update_many(
                                        {"user_id": user_id, "sonarStatus": "checked"},
                                        {"$set": {"all_task_email": "sent"}}
                                    )
                                    print("Пользователь завершил задания (админу)")
            # users = collection.distinct("user_id", {"all_task_email": {"$ne": "sent"}})
            #
            # for user_id in users:
            #     try:
            #         user_tasks = list(collection.find({"user_id": user_id}))
            #         all_checked = all(task['sonarStatus'] == "checked" for task in user_tasks)
            #
            #         if all_checked and user_tasks:
            #             user_document = users_collection.find_one({"_id": ObjectId(user_id)})
            #             if not user_document:
            #                 continue
            #
            #             surname = user_document['surname']
            #             name = user_document['name']
            #             patro = user_document.get('patro', '')
            #
            #             task_ids = [task['task_id'] for task in user_tasks]
            #             task_documents = [tasks_collection.find_one({"_id": ObjectId(task_id)}) for task_id in task_ids]
            #
            #             task_numbers_sorted = sorted(
            #                 (task_doc['taskNumber'] for task_doc in task_documents if task_doc),
            #                 key=int
            #             )
            #
            #             task_documents_sorted = {
            #                 task_doc['taskNumber']: task_doc for task_doc in task_documents if task_doc
            #             }
            #
            #             message_text = (
            #                 f"Пользователь {surname} {name} {patro} завершил все выданные Вами тестовые задания.\n"
            #                 f"Назначьте новое задание для продолжения тестирования, или свяжитесь с кандидатом.\n"
            #                 f"С выгрузкой его успехов можете ознакомиться в прикрепленном файле."
            #             )
            #
            #             admin_emails = admins_collection.distinct("email")
            #
            #             for admin_email in admin_emails:
            #                 for task_number in task_numbers_sorted:
            #                     task_document = task_documents_sorted.get(task_number)
            #                     if task_document:
            #                         user_task_document = next(
            #                             (ut for ut in user_tasks if ut['task_id'] == str(task_document['_id'])),
            #                             None
            #                         )
            #                         if user_task_document:
            #                             save_and_send_report(user_document, task_document, user_task_document,
            #                                                  admin_email)
            #
            #                 if send_email(admin_email, 'Пользователь завершил задания', message_text, "Admin"):
            #                     collection.update_many(
            #                         {"user_id": user_id, "sonarStatus": "checked"},
            #                         {"$set": {"all_task_email": "sent"}}
            #                     )
            #     except Exception as e:
            #         print(f"Exception occurred while processing documentssss: {str(e)}")
        except Exception as e:
            print(f"Exception occurred while processing documents: {str(e)}")

        time.sleep(10)  # Проверка каждые 10 секунд

if __name__ == '__main__':
    check_and_send_email()



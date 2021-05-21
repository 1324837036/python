import shutil
from MyApp.models import *
from django.http import JsonResponse, HttpResponse
import json
from django.core.mail import send_mail, send_mass_mail
from django.contrib.auth.hashers import make_password, check_password
import xlwt
import xlrd
import io
from django.utils.http import urlquote
import re
from django.forms.models import model_to_dict
import os
import requests
from django.core.files.storage import default_storage
from urllib.parse import quote
from xlutils.copy import copy
from MyLibrary import settings
import time

def test(request):
    return JsonResponse({"status": 0, "message": "This is Django message"})


def add_Wait_persons(request):
    file1 = request.FILES.get("file", None)
    managerId = request.POST.get('managerId')
    if os.path.exists('成员信息.xlsx'):
        os.remove('成员信息.xlsx')
    default_storage.save('成员信息.xlsx', file1)
    work_book = xlrd.open_workbook('成员信息.xlsx')
    sheet_1 = work_book.sheet_by_index(0)
    if (sheet_1.cell(0, 0).value != '工号' or sheet_1.cell(0, 1).value != '姓名' or
            sheet_1.cell(0, 2).value != '部门' or sheet_1.cell(0, 3).value != '邮箱' or
            sheet_1.cell(0, 4).value != '职称' or sheet_1.cell(0, 5).value != '学校'):
        return JsonResponse({"state": 0, "message": "文件中存在不正确的信息，请修改后重新提交"})

    for i in range(1, sheet_1.nrows):
        if (sheet_1.cell(i, 0).ctype == 0 or sheet_1.cell(i, 1).ctype == 0 or
                sheet_1.cell(i, 2).ctype == 0 or sheet_1.cell(i, 3).ctype == 0 or
                sheet_1.cell(i, 4).ctype == 0 or sheet_1.cell(i, 5).ctype == 0 or
                sheet_1.cell(i, 0).value == '' or sheet_1.cell(i, 1).value == '' or
                sheet_1.cell(i, 2).value == '' or sheet_1.cell(i, 3).value == '' or
                sheet_1.cell(i, 4).value == '' or sheet_1.cell(i, 5).value == ''):
            return JsonResponse({"state": 0, "message": "文件中存在不正确的信息，请修改后重新提交"})
    for i in range(1, sheet_1.nrows):
        if len(Wait_persons.objects.filter(id1=sheet_1.cell(i, 0).value,managerId=managerId)) == 0:
            Wait_persons.objects.create(
                id1=sheet_1.cell(i, 0).value, name=sheet_1.cell(i, 1).value,
                department=sheet_1.cell(i, 2).value, email=sheet_1.cell(i, 3).value,
                title=sheet_1.cell(i, 4).value, orgName=sheet_1.cell(i, 5).value,
                managerId=managerId)
    return JsonResponse({"state": 1, "message": "文件已经添加进待导入成员数据库"},
                        json_dumps_params={"ensure_ascii": False})


def get_Wait_persons(request):
    managerId = request.POST.get('managerId')
    data = []
    for person in Wait_persons.objects.filter(managerId=managerId):
        dic2 = model_to_dict(person)
        dic2['id'] = dic2['id1']
        data.append(dic2)
    data1 = dict()
    data1['data'] = data
    data1['message'] = 'success'
    data1['confirm_checked_num'] = len(Imported_persons.objects.filter(managerId=managerId))
    return JsonResponse(data1, safe=False, json_dumps_params={"ensure_ascii": False})


def add_Imported_persons(request):
    data = json.loads(request.POST.get('data'))
    managerId = request.POST.get('managerId')
    if len(Imported_persons.objects.filter(id1=data['id'], managerId=managerId)) == 0:
        Imported_persons.objects.create(
            id1=data['id'],
            name=data['name'],
            department=data['department'],
            email=data['email'],
            title=data['title'],
            orgName=data['orgName'],
            avg=data['avg'],
            paperCount=data['paperCount'],
            projectCount=data['projectCount'],
            patentCount=data['patentCount'],
            awardCount=data['awardCount'],
            student_awardCount=data['student_awardCount'],
            workCount=data['workCount'],
            copyrightCount=data['copyrightCount'],
            scholarId=str(data['scholarId']),
            managerId=managerId
        )
    print(data['id'],managerId)
    Wait_persons.objects.filter(id1=data['id'], managerId=managerId).delete()
    return JsonResponse({
        "data": len(Imported_persons.objects.filter(managerId=managerId)),
        "status": 1, "message": "相关已经添加进导入成员数据库"}, json_dumps_params={"ensure_ascii": False})


def get_Imported_persons(request):
    managerId = request.POST.get('managerId')
    data1 = dict()
    data = []
    for person in Imported_persons.objects.filter(managerId=managerId):
        dic2 = model_to_dict(person)
        dic2['id'] = dic2['id1']
        data.append(dic2)
    data1['data'] = data
    data1['wait_to_confirm_num'] = len(Wait_persons.objects.filter(managerId=managerId))
    return JsonResponse(data1, safe=False, json_dumps_params={"ensure_ascii": False})


def remove_Imported_persons(request):
    id = request.POST.get('id')
    managerId = request.POST.get('managerId')
    Imported_persons.objects.filter(id1=id, managerId=managerId).delete()
    return JsonResponse({"message": "成员已成功删除"}, json_dumps_params={"ensure_ascii": False})


def change_Imported_persons(request):
    data = json.loads(request.POST.get('data'))
    managerId = request.POST.get('managerId')
    object1 = Imported_persons.objects.filter(id1=data['id'], managerId=managerId)
    object1.update(
        id=data['id'],
        name=data['name'],
        department=data['department'],
        email=data['email'],
        title=data['title'],
        orgName=data['orgName'],
        avg=data['avg'],
        paperCount=data['paperCount'],
        projectCount=data['projectCount'],
        patentCount=data['patentCount'],
        awardCount=data['awardCount'],
        student_awardCount=data['student_awardCount'],
        workCount=data['workCount'],
        copyrightCount=data['copyrightCount']
    )
    return JsonResponse({"message": "成员信息已经成功修改"}, json_dumps_params={"ensure_ascii": False})


def search_Imported_persons(request):
    name = request.POST.get('name')
    managerId = request.POST.get('managerId')
    data1 = []
    for person in Imported_persons.objects.filter(name=name, managerId=managerId):
        data2 = dict()
        data2['name'] = person.name
        data2['id'] = person.id
        data2['department'] = person.department
        data2['email'] = person.email
        data2['title'] = person.title
        data2['orgName'] = person.orgName
        data2['avg'] = person.avg
        data2['paperCount'] = person.paperCount
        data2['projectCount'] = person.projectCount
        data2['patentCount'] = person.patentCount
        data2['awardCount'] = person.awardCount
        data2['student_awardCount'] = person.student_awardCount
        data2['workCount'] = person.workCount
        data2['copyrightCount'] = person.copyrightCount
        data2['scholarId'] = person.scholarId
        data1.append(data2)
    data = {
        "data": data1,
        "message": "成员查询完毕"
    }
    return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})


def send_emails(request):
    managerId = request.POST.get("managerId")
    Wait_persons.objects.filter(managerId=managerId).delete()
    msg = '计算机学院邀请您，您的成果管理系统账户待注册,请点击下方链接完成注册\n'
    for person in Imported_persons.objects.filter(managerId=managerId):
        send_mail(
            subject="您好," + person.name + ',计算机学院邀请您进行成果管理系统账户注册',
            message="您好," + person.name + "," + msg
                    + "http://localhost:8081/Invitation?scholarId="
                    + str(person.scholarId) + '&managerId=' + str(managerId),
            from_email=settings.EMAIL_HOST_USER,
            recipient_list=[person.email]  # 这里注意替换成自己的目的邮箱，不然就发到我的邮箱来了：）
        )
    return JsonResponse({'message': '测试邮件已发出请注意查收'}, json_dumps_params={"ensure_ascii": False})


# 单独发送邮件
def send_email_Single(request):
    data = json.loads(request.POST.get('data'))
    email = data['email']
    name = data['name']
    msg = '计算机学院邀请您，您的成果管理系统账户待注册,请点击下方链接完成注册\n'
    send_mail(
        subject="您好," + name + '，计算机学院邀请您进行成果管理系统账户注册',
        message="您好," + name + "," + msg
                + "http://localhost:8081/Invitation?scholarId="
                + str(data['scholarId']),
        from_email=settings.EMAIL_HOST_USER,
        recipient_list=[email]  # 这里注意替换成自己的目的邮箱，不然就发到我的邮箱来了：）
    )
    return JsonResponse({'message': '测试邮件已发出请注意查收'}, json_dumps_params={"ensure_ascii": False})


# 单独发送成果确认邮件
def send_Achivement_email_Single(request):
    data = json.loads(request.POST.get('data'))
    email = data['email']
    name = data['name']
    scholarId = data['scholarId']
    begin_year = data['begin_year']
    end_year = data['end_year']
    managerId = data['managerId']
    Achievement_renew(data)
    orgmanager = orgManager.objects.filter(managerId=managerId).first()
    msg = orgmanager.orgName + orgmanager.department + '邀请您，您的' + str(begin_year) + '年到' + str(end_year) + \
          '年成果待确认,请点击下方链接完成注册\n'
    send_mail(
        subject="成果确认",
        message="您好，" + name + "，" + msg
                + "http://localhost:8081/teacher_Achievement_analysis?managerId="
                + str(managerId) + "&scholarId="
                + str(scholarId) + '&begin_year='
                + str(begin_year) + '&end_year=' + str(end_year),
        from_email=settings.EMAIL_HOST_USER,
        recipient_list=[email]  # 这里注意替换成自己的目的邮箱，不然就发到我的邮箱来了：）
    )
    return JsonResponse({'message': '测试邮件已发出请注意查收'}, json_dumps_params={"ensure_ascii": False})


# 群发成果邮件
def send_Achivement_emails(request):
    data = json.loads(request.POST.get('data'))
    begin_year = data['begin_year']
    end_year = data['end_year']
    report_id = Achievement_report.objects.filter(begin_year=begin_year, end_year=end_year).first()
    Achievement_report_detail.objects.filter(report_id=report_id).update(state=0)
    msg = '计算机学院邀请您，您的' + str(begin_year) + '年到' + str(end_year) + \
          '年成果待确认,请点击下方链接完成注册\n'
    for person in Achievement_report_detail.objects.filter(report_id=report_id):
        send_mail(
            subject="您好，" + person.name + '，计算机学院邀请您，您的成果待确认',
            message="您好，" + person.name + "，" + msg
                    + "http://localhost:8081/teacher_Achievement_analysis?scholarId="
                    + str(person.scholarId) + '&begin_year=' + str(begin_year) +
                    '&end_year=' + str(end_year),
            from_email=settings.EMAIL_HOST_USER,
            recipient_list=[person.email]  # 这里注意替换成自己的目的邮箱，不然就发到我的邮箱来了：）
        )
    return JsonResponse({'message': '测试邮件已发出请注意查收'}, json_dumps_params={"ensure_ascii": False})


def Achievement_renew(data):
    scholarId = data['scholarId']
    begin_year = data['begin_year']
    end_year = data['end_year']
    message = data['message']
    managerId = data['managerId']
    report_id = Achievement_report.objects.filter(begin_year=begin_year,
                                                  end_year=end_year,
                                                  managerId=managerId).first()
    Achievement_report_detail.objects.filter(report_id=report_id,
                                             scholarId=scholarId,
                                             managerId=managerId).update(
        state=0,
        paperCount=message['paperCount'],
        paperSciCount=message['paperSciCount'],
        paperEiCount=message['paperEiCount'],
        paperOtherCount=message['paperOtherCount'],
        projectCount=message['projectCount'],
        projectNationCount=message['projectNationCount'],
        projectProvinceCount=message['projectProvinceCount'],
        projectOtherCount=message['projectOtherCount'],
        patentCount=message['patentCount'],
        awardCount=message['awardCount'],
        awardNationCount=message['awardNationCount'],
        awardProvinceCount=message['awardProvinceCount'],
        awardOtherCount=message['awardOtherCount'],
        student_awardCount=message['student_awardCount'],
        student_awardNationCount=message['student_awardNationCount'],
        student_awardProvinceCount=message['student_awardProvinceCount'],
        student_awardOtherCount=message['student_awardOtherCount'],
        workCount=message['workCount'],
        software_copyrightCount=message['software_copyrightCount'],
    )


# 添加成果报告
def add_Achievement_report(request):
    data = json.loads(request.POST.get('data'))
    managerId = request.POST.get('managerId')
    if len(Achievement_report.objects.filter(
            begin_year=data['begin_year'],
            end_year=data['end_year'],
            managerId=managerId)) == 0:

        Achievement_report.objects.create(
            begin_year=data['begin_year'],
            end_year=data['end_year'],
            managerId=managerId
        )
        data = {
            'status': 1,
            'report_id': Achievement_report.objects.filter(
                begin_year=data['begin_year'],
                end_year=data['end_year'],
                managerId=managerId
            )[0].report_id,
            'messages': '报告添加完成'
        }
    else:
        data = {
            'status': 0,
            'messages': '该报告已存在，请勿重复添加'
        }
    return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})


# 添加年度成果报告详情
def add_Achievement_report_detail(request):
    list = json.loads(request.POST.get("data"))
    managerId = request.POST.get("managerId")
    for i in range(len(list)):
        data = list[i]
        Achievement_report_detail.objects.create(
            report_id=Achievement_report.objects.filter(report_id=data['report_id'])[0],
            state=data['state'],
            id1=str(data['id']),
            name=data['name'],
            email=data['email'],
            scholarId=data['scholarId'],
            paperCount=data['paperCount'],
            paperSciCount=data['paperSciCount'],
            paperEiCount=data['paperEiCount'],
            paperOtherCount=data['paperOtherCount'],
            projectCount=data['projectCount'],
            projectNationCount=data['projectNationCount'],
            projectProvinceCount=data['projectProvinceCount'],
            projectOtherCount=data['projectOtherCount'],
            patentCount=data['patentCount'],
            awardCount=data['awardCount'],
            awardNationCount=data['awardNationCount'],
            awardProvinceCount=data['awardProvinceCount'],
            awardOtherCount=data['awardOtherCount'],
            student_awardCount=data['student_awardCount'],
            student_awardNationCount=data['student_awardNationCount'],
            student_awardProvinceCount=data['student_awardProvinceCount'],
            student_awardOtherCount=data['student_awardOtherCount'],
            workCount=data['workCount'],
            software_copyrightCount=data['software_copyrightCount'],
            managerId=managerId
        )
    data = {
        'status': 1,
        'messages': '信息添加完成'
    }
    return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})


def get_Achievement_report(request):
    managerId = request.POST.get('managerId')
    data = []
    for report in Achievement_report.objects.filter(managerId=managerId):
        data2 = dict()
        data2['begin_year'] = report.begin_year
        data2['end_year'] = report.end_year
        data.append(data2)
    data1 = {
        'data': data,
        'messages': '数据返回成功',
        'state': 1
    }
    return JsonResponse(data1, safe=False, json_dumps_params={"ensure_ascii": False})


def get_Achievement_report_detail(request):
    begin_year = request.GET.get("begin_year")
    end_year = request.GET.get("end_year")
    managerId = request.GET.get("managerId")
    report_id = Achievement_report.objects.filter(
        begin_year=begin_year,
        end_year=end_year,
        managerId=managerId)[0]
    report_details = Achievement_report_detail.objects.filter(report_id=report_id)
    list1 = []
    for report_detail in report_details:
        dic = model_to_dict(report_detail)
        dic['state'] = '未确认' if (report_detail.state == 0) else '已确认'
        dic['total'] = report_detail.paperCount + report_detail.projectCount \
                       + report_detail.projectCount + report_detail.awardCount \
                       + report_detail.student_awardCount + report_detail.software_copyrightCount
        list1.append(dic)
    data1 = {
        'data': list1,
        'messages': '数据返回成功',
        'status': 1
    }
    return JsonResponse(data1, safe=False, json_dumps_params={"ensure_ascii": False})


def get_Excel(request):
    list = json.loads(request.POST.get("data"))
    name = urlquote('相关文件导出表')
    work_book = xlwt.Workbook(encoding='utf-8')
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    style.alignment = al
    sheet = work_book.add_sheet('sheet')
    fifth_col = sheet.col(5)
    senventh_col = sheet.col(7)
    fifth_col.width = 150 * 25
    senventh_col.width = 150 * 25
    sheet.write(0, 0, '姓名', style)
    sheet.write(0, 1, '论文', style)
    sheet.write(0, 2, '项目', style)
    sheet.write(0, 3, '专利', style)
    sheet.write(0, 4, '获奖', style)
    sheet.write(0, 5, '软件著作权', style)
    sheet.write(0, 6, '学生获奖', style)
    sheet.write(0, 7, '总数', style)
    for i in range(len(list)):
        person = list[i]
        sheet.write(i + 1, 0, person['name'], style)
        sheet.write(i + 1, 1, person['paperCount'], style)
        sheet.write(i + 1, 2, person['projectCount'], style)
        sheet.write(i + 1, 3, person['patentCount'], style)
        sheet.write(i + 1, 4, person['awardCount'], style)
        sheet.write(i + 1, 5, person['software_copyrightCount'], style)
        sheet.write(i + 1, 6, person['student_awardCount'], style)
        sheet.write(i + 1, 7, person['total'], style)
    output = io.BytesIO()
    work_book.save(output)
    response = HttpResponse(content_type="application/vnd.ms-excel")
    response['Content-Disposition'] = 'attachment;filename={0}.xlsx'.format(quote(name))
    response.write(output.getvalue())
    return response


def get_Excel2(request):
    list = json.loads(request.POST.get("data"))
    name = urlquote('相关文件导出表')
    work_book = xlwt.Workbook(encoding='utf-8')
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    style.alignment = al
    sheet = work_book.add_sheet('sheet')
    fifth_col = sheet.col(5)
    senventh_col = sheet.col(7)
    tenth_col = sheet.col(10)
    eleventh_col = sheet.col(11)
    fourteenth_col = sheet.col(14)
    fifteenth_col = sheet.col(15)
    sixteenth_col = sheet.col(16)
    seventeenth_col = sheet.col(17)
    eighteenth_col = sheet.col(18)

    fifth_col.width = 150 * 25
    senventh_col.width = 150 * 25
    tenth_col.width = 150 * 25
    eleventh_col.width = 150 * 25
    fourteenth_col.width = 150 * 25
    fifteenth_col.width = 150 * 25
    sixteenth_col.width = 150 * 25
    seventeenth_col.width = 150 * 25
    eighteenth_col.width = 150 * 25

    sheet.write(0, 0, '姓名', style)
    sheet.write(0, 1, 'SCI论文', style)
    sheet.write(0, 2, 'EI论文', style)
    sheet.write(0, 3, '其他论文', style)
    sheet.write(0, 4, '论文总数', style)
    sheet.write(0, 5, '国家级项目', style)
    sheet.write(0, 6, '省部级项目', style)
    sheet.write(0, 7, '其他项目', style)
    sheet.write(0, 8, '项目总数', style)
    sheet.write(0, 9, '专利', style)
    sheet.write(0, 10, '国家级获奖', style)
    sheet.write(0, 11, '省部级获奖', style)
    sheet.write(0, 12, '其他获奖', style)
    sheet.write(0, 13, '获奖总数', style)
    sheet.write(0, 14, '软件著作权', style)
    sheet.write(0, 15, '国家级学生获奖', style)
    sheet.write(0, 16, '省部级学生获奖', style)
    sheet.write(0, 17, '其他学生获奖', style)
    sheet.write(0, 18, '学生获奖总数', style)
    sheet.write(0, 19, '成果总数', style)
    sheet.write(0, 20, '确认状态', style)
    for i in range(len(list)):
        person = list[i]
        sheet.write(i + 1, 0, person['name'], style)
        sheet.write(i + 1, 1, person['paperSciCount'], style)
        sheet.write(i + 1, 2, person['paperEiCount'], style)
        sheet.write(i + 1, 3, person['paperOtherCount'], style)
        sheet.write(i + 1, 4, person['paperCount'], style)
        sheet.write(i + 1, 5, person['projectNationCount'], style)
        sheet.write(i + 1, 6, person['projectProvinceCount'], style)
        sheet.write(i + 1, 7, person['projectOtherCount'], style)
        sheet.write(i + 1, 8, person['projectCount'], style)
        sheet.write(i + 1, 9, person['patentCount'], style)
        sheet.write(i + 1, 10, person['awardNationCount'], style)
        sheet.write(i + 1, 11, person['awardProvinceCount'], style)
        sheet.write(i + 1, 12, person['awardOtherCount'], style)
        sheet.write(i + 1, 13, person['awardCount'], style)
        sheet.write(i + 1, 14, person['software_copyrightCount'], style)
        sheet.write(i + 1, 15, person['student_awardNationCount'], style)
        sheet.write(i + 1, 16, person['student_awardProvinceCount'], style)
        sheet.write(i + 1, 17, person['student_awardOtherCount'], style)
        sheet.write(i + 1, 18, person['student_awardCount'], style)
        sheet.write(i + 1, 19, person['total'], style)
        sheet.write(i + 1, 20, person['state'], style)
    output = io.BytesIO()
    work_book.save(output)
    response = HttpResponse(content_type="application/vnd.ms-excel")
    response['Content-Disposition'] = 'attachment;filename={0}.xlsx'.format(quote(name))
    response.write(output.getvalue())
    return response


# 邀请注册
def add_Admin_messages(request):
    data1 = json.loads(request.POST.get("data"))
    scholarId = data1['scholarId']
    managerId = request.POST.get("managerId")
    p = re.compile('^[A-Za-z0-9\u4e00-\u9fa5]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$')
    x = p.match(data1['email'])
    if x is None:
        data = {
            'state': 0,
            'messages': '邮箱格式输入不争取'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})
    if data1['password'] != data1['confirm_password']:
        data = {
            'state': 0,
            'messages': '两次密码输入不同，请重新输入'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})

    if len(Imported_persons.objects.filter(
            scholarId=scholarId, managerId=managerId)) == 0:
        data = {
            'state': 0,
            'messages': '您尚未被添加到系统中，请联系管理人员进行添加'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})

    if len(Admin_messages.objects.filter(email=data1['email'])) != 0:
        data = {
            'state': 0,
            'messages': '该邮箱已被注册，请联系系统管理人员更换邮箱'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})

    if data1['email'] != Imported_persons.objects.filter(
            scholarId=scholarId, managerId=managerId).first().email:
        data = {
            'state': 0,
            'messages': '该邮箱与您上报的邮箱不符，请联系管理人员进行邮箱更换'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})
    Admin_messages.objects.create(
        email=data1['email'],
        password=make_password(data1['password']),
        scholarId=scholarId
    )
    scholarId = data1['scholarId']
    projects = ['论文', '专利', '项目', '软件著作权', '获奖', '学生获奖']
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    for project in projects:
        path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/' + project)
        if not os.path.exists(path):
            os.makedirs(path)
    data = {
        'state': 1,
        'messages': '账号注册成功'
    }

    return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})


# 登录
def login(request):
    data1 = json.loads(request.POST.get("data"))
    p = re.compile('^[A-Za-z0-9\u4e00-\u9fa5]+@[a-zA-Z0-9_-]+(\.[a-zA-Z0-9_-]+)+$')
    x = p.match(data1['email'])
    if x is None:
        data = {
            'status': 0,
            'messages': '邮箱格式输入不争取'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})
    object = Admin_messages.objects.filter(
        email=data1['email'])
    if len(object) == 0:
        data = {
            'state': 0,
            'messages': '邮箱账号不存在'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})
    if not check_password(data1['password'], object[0].password):
        data = {
            'state': 0,
            'messages': '密码错误'
        }
        return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})
    data = {
        'state': 1,
        'person_state': object[0].state,
        'messages': '登录成功',
        'scholarId': object[0].scholarId
    }
    return JsonResponse(data, safe=False, json_dumps_params={"ensure_ascii": False})


# 通过学者id获取某一学者基本信息
def get_messageByScholarId(request):
    scholarId = request.POST.get("scholarId")
    message = Imported_persons.objects.filter(scholarId=scholarId).first()
    if len(AchievementCount.objects.filter(scholarId=scholarId)) == 0:
        AchievementCount.objects.create(scholarId=scholarId)
    achievementCount = AchievementCount.objects.filter(scholarId=scholarId).first()
    data1 = {}
    data = model_to_dict(message)
    data['paperCount_unchecked'] = message.paperCount - achievementCount.paperCount
    data['patentCount_unchecked'] = message.patentCount - achievementCount.patentCount
    data['projectCount_unchecked'] = message.projectCount - achievementCount.projectCount
    data['awardCount_unchecked'] = message.awardCount - achievementCount.awardCount
    data['workCount_unchecked'] = message.workCount - achievementCount.workCount
    data['student_award_Count_unchecked'] = message.student_awardCount - achievementCount.studentAwardCount
    data['copyrightCount_unchecked'] = message.copyrightCount - achievementCount.copyrightCount
    data['total_number'] = data['paperCount_unchecked'] + data["patentCount_unchecked"] + \
                           data['projectCount_unchecked'] + \
                           data['awardCount_unchecked'] + data['workCount_unchecked'] + \
                           data['student_award_Count_unchecked'] + \
                           data['copyrightCount_unchecked']

    data['paperCount_checked'] = achievementCount.paperCount
    data['patentCount_checked'] = achievementCount.patentCount
    data['projectCount_checked'] = achievementCount.projectCount
    data['awardCount_checked'] = achievementCount.awardCount
    data['workCount_checked'] = achievementCount.workCount
    data['student_award_Count_checked'] = achievementCount.studentAwardCount
    data['copyrightCount_checked'] = achievementCount.copyrightCount
    data['total_number_checked'] = data['paperCount_checked'] + data["patentCount_checked"] + \
                                   data['projectCount_checked'] + \
                                   data['awardCount_checked'] + data['workCount_checked'] + \
                                   data['student_award_Count_checked'] + \
                                   data['copyrightCount_checked']

    data['paperCount2'] = data['paperCount'] + len(PaperMessage.objects.filter(scholarId=scholarId))
    data['patentCount2'] = data['patentCount'] + len(PatentMessage.objects.filter(scholarId=scholarId))
    data['projectCount2'] = data['projectCount'] + \
                            len(ProjectTransverseMessage.objects.filter(scholarId=scholarId)) + \
                            len(ProjectVerticalMessage.objects.filter(scholarId=scholarId))
    data['awardCount2'] = data['awardCount'] + len(AwardMessage.objects.filter(scholarId=scholarId))
    data['student_awardCount2'] = data['student_awardCount'] + len(
        StudentAwardMessage.objects.filter(scholarId=scholarId))
    data['copyrightCount2'] = data['copyrightCount'] + len(SoftwareCopyrightMessage.objects.filter(scholarId=scholarId))
    data1['data'] = data
    data1['state'] = 200
    return JsonResponse(data1, safe=False, json_dumps_params={"ensure_ascii": False})


# 获取文件夹下的所有文件数
def get_FolderNum(request):
    scholarId = request.POST.get('scholarId')
    data = model_to_dict(Imported_persons.objects.filter(scholarId=scholarId).first())
    projects = ['论文/已导入论文信息.xls', '专利/已导入专利信息.xls', '项目/已导入项目信息.xls', '软件著作权/已导入软件著作权信息.xls',
                '获奖/已导入获奖信息.xls', '学生获奖/已导入学生获奖信息.xls']
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId)
    list1 = []
    for project_name in projects:
        if (os.path.exists(os.path.join(path, project_name))):
            work_book = xlrd.open_workbook(os.path.join(path, project_name))
            table = work_book.sheets()[0]
            nrows_ori = table.nrows
            list1.append(nrows_ori - 1)
        else:
            list1.append(0)
    data['paperFolderNum'] = list1[0]
    data['patentFolderNum'] = list1[1]
    data['projectFolderNum'] = list1[2]
    data['copyrightFolderNum'] = list1[3]
    data['awardFolderNum'] = list1[4]
    data['student_awardFolderNum'] = list1[5]
    data['workFolderNum'] = 0
    data['totalFolderNum'] = list1[0] + list1[1] + list1[2] + list1[3] + list1[4] + list1[5]
    data1 = {
        'state': 1,
        'message': '获取文件夹数完成',
        'data': data
    }
    return JsonResponse(data1, safe=False, json_dumps_params={"ensure_ascii": False})


# 保存文件
def save_file(request):
    file = request.FILES.get("file", None)
    project_name = request.POST.get("project_name")
    scholarId = request.POST.get("scholarId")
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/' + project_name)
    if not os.path.exists(path):
        os.makedirs(path)
    path = os.path.join(path, file.name.replace('.xlsx', '.xls'))
    default_storage.save(path, file)
    if judge_format(path,project_name):
        if judge_repeat(path, os.path.join(
                BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' +
                          scholarId + '/' + project_name + '/已导入' + project_name + '信息.xls'), project_name, scholarId):
            try:
                if project_name == '论文':
                    save_PaperMessage(path, scholarId)
                    os.remove(path)
                elif project_name == '专利':
                    save_PatentMessage(path, scholarId)
                    os.remove(path)
                elif project_name == '项目':
                    save_ProjectMessage(path, scholarId)
                    os.remove(path)
                elif project_name == '软件著作权':
                    save_SoftwareCopyrightMessage(path, scholarId)
                    os.remove(path)
                elif project_name == '获奖':
                    save_AwardMessage(path, scholarId)
                    os.remove(path)
                elif project_name == '学生获奖':
                    save_StudentAwardMessage(path, scholarId)
                    os.remove(path)
            except:
                os.remove(path)
                return JsonResponse({"state": 2, "message": "文件中数据格式不正确，请重新上传"}, json_dumps_params={"ensure_ascii": False})
            return JsonResponse({"state": 1, "message": "上传文件成功"}, json_dumps_params={"ensure_ascii": False})
        else:
            os.remove(path)
            return JsonResponse({"state": 0, "message": "出现重名成果名字，请更改名字后重新上传"}, json_dumps_params={"ensure_ascii": False})
    else:
        os.remove(path)
        return JsonResponse({"state": 2, "message": "文件中数据格式不正确，请重新上传"}, json_dumps_params={"ensure_ascii": False})

# 判断文件中数据格式是否有误
def judge_format(path1, project_name):
    work_book = xlrd.open_workbook(path1)
    sheet_1 = work_book.sheet_by_index(0)
    table = work_book.sheets()[0]
    nrows_ori = table.nrows
    if project_name == '论文':
        for i in range(1, nrows_ori):
            try:
                year = int(sheet_1.cell(i, 2).value)
                issue = int(sheet_1.cell(i, 8).value)
                volume = int(sheet_1.cell(i, 9).value)
                pageStart = int(sheet_1.cell(i, 10).value)
                pageEnd = int(sheet_1.cell(i, 11).value)
                citationNum = int(sheet_1.cell(i, 13).value)
                if sheet_1.cell(i, 16).value not in ['是', '否'] or \
                        sheet_1.cell(i, 17).value not in ['是', '否']:
                    return False
            except:
                return False
    elif project_name == '专利':
        for i in range(1, nrows_ori):
            if sheet_1.cell(i, 2).value not in ['在审', '有效']:
                return False
    elif project_name == '项目':
        if (sheet_1.cell(0, 1).value == '合同来源'):
            try:
                for i in range(1, nrows_ori):
                    if sheet_1.cell(i, 3).value not in ['国家级项目', '省部级项目', '其他项目']:
                        return False
                    fund = float(sheet_1.cell(i, 5).value),
                    startYear = int(sheet_1.cell(i, 6).value),
                    endYear = int(sheet_1.cell(i, 7).value),
                    if sheet_1.cell(i, 8).value not in ['正在进行', '已结束']:
                        return False
            except:
                return False
        else:
            try:
                for i in range(1, nrows_ori):
                    if sheet_1.cell(i, 3).value not in ['国家级项目', '省部级项目', '其他项目']:
                        return False
                    fund = sheet_1.cell(i, 5).value
                    startYear = int(sheet_1.cell(i, 6).value)
                    endYear = int(sheet_1.cell(i, 7).value)
                    if sheet_1.cell(i, 8).value not in ['正在进行', '已结束']:
                        return False
            except:
                return False
    elif project_name == '软件著作权':
        try:
            for i in range(1, nrows_ori):
                endTime = int(sheet_1.cell(i, 2).value)
                getTime = int(sheet_1.cell(i, 3).value)
        except:
            return False
    elif project_name == '获奖':
        try:
            for i in range(1, nrows_ori):
                if sheet_1.cell(i, 2).value not in ['国家级获奖', '省部级获奖', '其他获奖']:
                    return False
                getTime = int(sheet_1.cell(i, 4).value),
        except:
            return False
    elif project_name == '学生获奖':
        try:
            for i in range(1, nrows_ori):
                if sheet_1.cell(i, 3).value not in ['国家级获奖', '省部级获奖', '其他获奖']:
                    return False
                getTime = int(sheet_1.cell(i, 5).value),
        except:
            return False
    return True


# 判断文件是否有重复信息
def judge_repeat(path1, path2, project_name, scholarId):
    if project_name in ['论文', '专利', '项目', '软件著作权']:
        if (os.path.exists(path2)):
            work_book = xlrd.open_workbook(path2)
            sheet_1 = work_book.sheet_by_index(0)
            col_0_value = sheet_1.col_values(0)
            del col_0_value[0]
            work_book2 = xlrd.open_workbook(path1)
            sheet_1_2 = work_book2.sheet_by_index(0)
            col_0_value2 = sheet_1_2.col_values(0)
            del col_0_value2[0]
            set_list = set(col_0_value2)
            if len(set_list) != len(col_0_value2):
                return False
            ret = list(set(col_0_value) & set(col_0_value2))
            return len(ret) == 0
        else:
            return True
    else:
        if project_name == '获奖':
            work_book = xlrd.open_workbook(path1)
            sheet_1 = work_book.sheet_by_index(0)
            table = work_book.sheets()[0]
            nrows_ori = table.nrows
            list_keep = []
            for i in range(1, nrows_ori):
                if 0 != len(AwardMessage.objects.filter(
                        scholarId=scholarId,
                        title=sheet_1.cell(i, 0).value,
                        rank=sheet_1.cell(i, 1).value,
                        level=sheet_1.cell(i, 2).value,
                        org=sheet_1.cell(i, 3).value,
                        getTime=int(sheet_1.cell(i, 4).value),
                        topics=sheet_1.cell(i, 5).value,
                        authors=sheet_1.cell(i, 6).value,
                )):
                    return False
                str1 = sheet_1.cell(i, 0).value+sheet_1.cell(i, 1).value+\
                            sheet_1.cell(i, 2).value+sheet_1.cell(i, 3).value+\
                            str (sheet_1.cell(i, 4).value) + sheet_1.cell(i, 5).value+\
                            sheet_1.cell(i, 6).value
                if str1 in list_keep:
                    return False
                list_keep.append(str1)

        elif project_name == '学生获奖':
            work_book = xlrd.open_workbook(path1)
            sheet_1 = work_book.sheet_by_index(0)
            table = work_book.sheets()[0]
            nrows_ori = table.nrows
            list_keep = []
            for i in range(1, nrows_ori):
                if 0 != len(StudentAwardMessage.objects.filter(
                        scholarId=scholarId,
                        title=sheet_1.cell(i, 0).value,
                        student=sheet_1.cell(i, 1).value,
                        rank=sheet_1.cell(i, 2).value,
                        level=sheet_1.cell(i, 3).value,
                        org=sheet_1.cell(i, 4).value,
                        getTime=int(sheet_1.cell(i, 5).value),
                        topics=sheet_1.cell(i, 6).value,
                        authors=sheet_1.cell(i, 7).value,
                )):
                    return False
                str1 = sheet_1.cell(i, 0).value + sheet_1.cell(i, 1).value + \
                      sheet_1.cell(i, 2).value + sheet_1.cell(i, 3).value + \
                      sheet_1.cell(i, 4).value + str(sheet_1.cell(i, 5).value) + \
                      sheet_1.cell(i, 6).value + sheet_1.cell(i, 7).value
                if str1 in list_keep:
                    return False
                list_keep.append(str1)
        return True


# 保存论文信息
def save_PaperMessage(path, scholarId):
    work_book = xlrd.open_workbook(path)
    sheet_1 = work_book.sheet_by_index(0)
    table = work_book.sheets()[0]
    nrows_ori = table.nrows
    for i in range(1, nrows_ori):
        sciPaper = True if sheet_1.cell(i, 16).value == '是' else False
        eiPaper = True if sheet_1.cell(i, 17).value == '是' else False
        PaperMessage.objects.create(
            scholarId=scholarId,
            title=sheet_1.cell(i, 0).value,
            abst=sheet_1.cell(i, 1).value,
            year=int(sheet_1.cell(i, 2).value),
            docType=sheet_1.cell(i, 3).value,
            doi=sheet_1.cell(i, 4).value,
            lang=sheet_1.cell(i, 5).value,
            venue=sheet_1.cell(i, 6).value,
            filed=sheet_1.cell(i, 14).value,
            issue=int(sheet_1.cell(i, 8).value),
            volume=int(sheet_1.cell(i, 9).value),
            pageStart=int(sheet_1.cell(i, 10).value),
            pageEnd=int(sheet_1.cell(i, 11).value),
            publisher=sheet_1.cell(i, 12).value,
            citationNum=int(sheet_1.cell(i, 13).value),
            keywords=sheet_1.cell(i, 14).value,
            authors=sheet_1.cell(i, 15).value,
            sciPaper=sciPaper,
            eiPaper=eiPaper
        )
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02
    style.alignment = al
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/论文')
    path = os.path.join(path, '已导入论文信息.xls')
    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '论文名', style)
        sheet.write(0, 1, '摘要', style)
        sheet.write(0, 2, '发表年份', style)
        sheet.write(0, 3, '论文类型', style)
        sheet.write(0, 4, 'DOI号', style)
        sheet.write(0, 5, '语言', style)
        sheet.write(0, 6, '刊物名称', style)
        sheet.write(0, 7, '论文领域', style)
        sheet.write(0, 8, '议题', style)
        sheet.write(0, 9, '卷号', style)
        sheet.write(0, 10, '起始页码', style)
        sheet.write(0, 11, '结束页码', style)
        sheet.write(0, 12, '出版社', style)
        sheet.write(0, 13, '引用次数', style)
        sheet.write(0, 14, '关键词', style)
        sheet.write(0, 15, '作者信息', style)
        sheet.write(0, 16, '是否为sci', style)
        sheet.write(0, 17, '是否为ei', style)
        work_book.save(path)

    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    table = work_book.sheets()[0]
    nrows = table.nrows
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    fifth_col = sheet.col(4)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    fifteenth_col = sheet.col(14)
    sixteenth_col = sheet.col(15)
    first_col.width = 256 * 40
    sec_col.width = 256 * 80
    fifth_col.width = 256 * 25
    seventh_col.width = 256 * 25
    eightth_col.width = 256 * 25
    fifteenth_col.width = 256 * 25
    sixteenth_col.width = 256 * 25
    for i in range(1, nrows_ori):
        sheet.write(nrows, 0, sheet_1.cell(i, 0).value)
        sheet.write(nrows, 1, sheet_1.cell(i, 1).value)
        sheet.write(nrows, 2, sheet_1.cell(i, 2).value)
        sheet.write(nrows, 3, sheet_1.cell(i, 3).value)
        sheet.write(nrows, 4, sheet_1.cell(i, 4).value)
        sheet.write(nrows, 5, sheet_1.cell(i, 5).value)
        sheet.write(nrows, 6, sheet_1.cell(i, 6).value)
        sheet.write(nrows, 7, sheet_1.cell(i, 7).value)
        sheet.write(nrows, 8, sheet_1.cell(i, 8).value)
        sheet.write(nrows, 9, sheet_1.cell(i, 9).value)
        sheet.write(nrows, 10, sheet_1.cell(i, 10).value)
        sheet.write(nrows, 11, sheet_1.cell(i, 11).value)
        sheet.write(nrows, 12, sheet_1.cell(i, 12).value)
        sheet.write(nrows, 13, sheet_1.cell(i, 13).value)
        sheet.write(nrows, 14, sheet_1.cell(i, 14).value)
        sheet.write(nrows, 15, sheet_1.cell(i, 15).value)
        sheet.write(nrows, 16, sheet_1.cell(i, 16).value)
        sheet.write(nrows, 17, sheet_1.cell(i, 17).value)
        nrows += 1
    excel.save(path)


# 保存专利信息
def save_PatentMessage(path, scholarId):
    work_book = xlrd.open_workbook(path)
    sheet_1 = work_book.sheet_by_index(0)
    table = work_book.sheets()[0]
    nrows_ori = table.nrows
    for i in range(1, nrows_ori):
        PatentMessage.objects.create(
            scholarId=scholarId,
            title=sheet_1.cell(i, 0).value,
            patentType=sheet_1.cell(i, 1).value,
            legalStatus=sheet_1.cell(i, 2).value,
            authorizationNum=sheet_1.cell(i, 3).value,
            inventorName=sheet_1.cell(i, 4).value,
            priorityDate=sheet_1.cell(i, 5).value,
            applicationNum=sheet_1.cell(i, 6).value,
            applicationDate=sheet_1.cell(i, 7).value,
        )
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    style.alignment = al
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/专利/已导入专利信息.xls')
    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '专利名称', style)
        sheet.write(0, 1, '专利类型', style)
        sheet.write(0, 2, '专利状态', style)
        sheet.write(0, 3, '专利编号', style)
        sheet.write(0, 4, '专利权人', style)
        sheet.write(0, 5, '授权公告日', style)
        sheet.write(0, 6, '申请编号', style)
        sheet.write(0, 7, '专利申请日', style)
        work_book.save(path)
    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    table = work_book.sheets()[0]
    nrows = table.nrows
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    third_col = sheet.col(2)
    fourth_col = sheet.col(3)
    fifth_col = sheet.col(4)
    sixth_col = sheet.col(5)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    first_col.width = 256 * 40
    sec_col.width = 256 * 25
    third_col.width = 256 * 25
    fourth_col.width = 256 * 25
    fifth_col.width = 256 * 25
    sixth_col.width = 256 * 25
    seventh_col.width = 256 * 25
    eightth_col.width = 256 * 25
    for i in range(1, nrows_ori):
        sheet.write(nrows, 0, sheet_1.cell(i, 0).value)
        sheet.write(nrows, 1, sheet_1.cell(i, 1).value)
        sheet.write(nrows, 2, sheet_1.cell(i, 2).value)
        sheet.write(nrows, 3, sheet_1.cell(i, 3).value)
        sheet.write(nrows, 4, sheet_1.cell(i, 4).value)
        sheet.write(nrows, 5, sheet_1.cell(i, 5).value)
        sheet.write(nrows, 6, sheet_1.cell(i, 6).value)
        sheet.write(nrows, 7, sheet_1.cell(i, 7).value)
        nrows += 1
    excel.save(path)


# 保存项目信息
def save_ProjectMessage(path, scholarId):
    work_book = xlrd.open_workbook(path)
    sheet_1 = work_book.sheet_by_index(0)
    table = work_book.sheets()[0]
    nrows_ori = table.nrows
    if (sheet_1.cell(0, 1).value == '合同来源'):
        for i in range(1, nrows_ori):
            ProjectTransverseMessage.objects.create(
                scholarId=scholarId,
                title=sheet_1.cell(i, 0).value,
                source=sheet_1.cell(i, 1).value,
                contractId=sheet_1.cell(i, 2).value,
                typeFirst=sheet_1.cell(i, 3).value,
                typeSecondary=sheet_1.cell(i, 4).value,
                fund=float(sheet_1.cell(i, 5).value),
                startYear=int(sheet_1.cell(i, 6).value),
                endYear=int(sheet_1.cell(i, 7).value),
                projectStatus=sheet_1.cell(i, 8).value,
                authors=sheet_1.cell(i, 9).value,
            )
        style = xlwt.easyxf('align: wrap on')
        al = xlwt.Alignment()
        al.horz = 0x02
        style.alignment = al
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/项目')
        path = os.path.join(path, '已导入项目信息.xls')
        if not os.path.exists(path):
            work_book = xlwt.Workbook()
            sheet = work_book.add_sheet('sheet')
            sheet.write(0, 0, '课题名称', style)
            sheet.write(0, 1, '合同来源', style)
            sheet.write(0, 2, '合同编号', style)
            sheet.write(0, 3, '基金名称', style)
            sheet.write(0, 4, '授予单位', style)
            sheet.write(0, 5, '课题级别（一级）', style)
            sheet.write(0, 6, '课题级别（二级）', style)
            sheet.write(0, 7, '到账经费（万）', style)
            sheet.write(0, 8, '开始时间', style)
            sheet.write(0, 9, '截止时间', style)
            sheet.write(0, 10, '项目状态', style)
            sheet.write(0, 11, '作者信息', style)
            work_book.save(path)
        work_book = xlrd.open_workbook(path)
        excel = copy(wb=work_book)
        sheet = excel.get_sheet(0)
        table = work_book.sheets()[0]
        nrows = table.nrows
        first_col = sheet.col(0)
        sec_col = sheet.col(1)
        third_col = sheet.col(2)
        fourth_col = sheet.col(3)
        fifth_col = sheet.col(4)
        sixth_col = sheet.col(5)
        seventh_col = sheet.col(6)
        eightth_col = sheet.col(7)
        first_col.width = 256 * 40
        sec_col.width = 256 * 25
        third_col.width = 256 * 25
        fourth_col.width = 256 * 25
        fifth_col.width = 256 * 25
        sixth_col.width = 256 * 25
        seventh_col.width = 256 * 25
        eightth_col.width = 256 * 25
        for i in range(1, nrows_ori):
            sheet.write(nrows, 0, sheet_1.cell(i, 0).value)
            sheet.write(nrows, 1, sheet_1.cell(i, 1).value)
            sheet.write(nrows, 2, sheet_1.cell(i, 2).value)
            sheet.write(nrows, 3, '无')
            sheet.write(nrows, 4, '无')
            sheet.write(nrows, 5, sheet_1.cell(i, 3).value)
            sheet.write(nrows, 6, sheet_1.cell(i, 4).value)
            sheet.write(nrows, 7, sheet_1.cell(i, 5).value)
            sheet.write(nrows, 8, sheet_1.cell(i, 6).value)
            sheet.write(nrows, 9, sheet_1.cell(i, 7).value)
            sheet.write(nrows, 10, sheet_1.cell(i, 8).value)
            sheet.write(nrows, 11, sheet_1.cell(i, 9).value)
            nrows += 1
        excel.save(path)
    else:
        for i in range(1, nrows_ori):
            ProjectVerticalMessage.objects.create(
                scholarId=scholarId,
                title=sheet_1.cell(i, 0).value,
                typeTertiary=sheet_1.cell(i, 1).value,
                org=sheet_1.cell(i, 2).value,
                typeFirst=sheet_1.cell(i, 3).value,
                typeSecondary=sheet_1.cell(i, 4).value,
                fund=sheet_1.cell(i, 5).value,
                startYear=int(sheet_1.cell(i, 6).value),
                endYear=int(sheet_1.cell(i, 7).value),
                projectStatus=sheet_1.cell(i, 8).value,
                authors=sheet_1.cell(i, 9).value,
            )
        style = xlwt.easyxf('align: wrap on')
        al = xlwt.Alignment()
        al.horz = 0x02
        style.alignment = al
        BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/项目')
        path = os.path.join(path, '已导入项目信息.xls')
        if not os.path.exists(path):
            work_book = xlwt.Workbook()
            sheet = work_book.add_sheet('sheet')
            sheet.write(0, 0, '课题名称', style)
            sheet.write(0, 1, '合同来源', style)
            sheet.write(0, 2, '合同编号', style)
            sheet.write(0, 3, '基金名称', style)
            sheet.write(0, 4, '授予单位', style)
            sheet.write(0, 5, '课题级别（一级）', style)
            sheet.write(0, 6, '课题级别（二级）', style)
            sheet.write(0, 7, '到账经费（万）', style)
            sheet.write(0, 8, '开始时间', style)
            sheet.write(0, 9, '截止时间', style)
            sheet.write(0, 10, '项目状态', style)
            sheet.write(0, 11, '作者信息', style)
            work_book.save(path)
        work_book = xlrd.open_workbook(path)
        excel = copy(wb=work_book)
        sheet = excel.get_sheet(0)
        table = work_book.sheets()[0]
        nrows = table.nrows
        first_col = sheet.col(0)
        sec_col = sheet.col(1)
        third_col = sheet.col(2)
        fourth_col = sheet.col(3)
        fifth_col = sheet.col(4)
        sixth_col = sheet.col(5)
        seventh_col = sheet.col(6)
        eightth_col = sheet.col(7)
        first_col.width = 256 * 40
        sec_col.width = 256 * 25
        third_col.width = 256 * 25
        fourth_col.width = 256 * 25
        fifth_col.width = 256 * 25
        sixth_col.width = 256 * 25
        seventh_col.width = 256 * 25
        eightth_col.width = 256 * 25
        for i in range(1, nrows_ori):
            sheet.write(nrows, 0, sheet_1.cell(i, 0).value)
            sheet.write(nrows, 1, '无')
            sheet.write(nrows, 2, '无')
            sheet.write(nrows, 3, sheet_1.cell(i, 1).value)
            sheet.write(nrows, 4, sheet_1.cell(i, 2).value)
            sheet.write(nrows, 5, sheet_1.cell(i, 3).value)
            sheet.write(nrows, 6, sheet_1.cell(i, 4).value)
            sheet.write(nrows, 7, sheet_1.cell(i, 5).value)
            sheet.write(nrows, 8, sheet_1.cell(i, 6).value)
            sheet.write(nrows, 9, sheet_1.cell(i, 7).value)
            sheet.write(nrows, 10, sheet_1.cell(i, 8).value)
            sheet.write(nrows, 11, sheet_1.cell(i, 9).value)
            nrows += 1
        excel.save(path)


# 保存软件著作权信息
def save_SoftwareCopyrightMessage(path, scholarId):
    work_book = xlrd.open_workbook(path)
    sheet_1 = work_book.sheet_by_index(0)
    table = work_book.sheets()[0]
    nrows_ori = table.nrows
    for i in range(1, nrows_ori):
        SoftwareCopyrightMessage.objects.create(
            scholarId=scholarId,
            title=sheet_1.cell(i, 0).value,
            certificateId=sheet_1.cell(i, 1).value,
            endTime=int(sheet_1.cell(i, 2).value),
            getTime=int(sheet_1.cell(i, 3).value),
            registrationNum=sheet_1.cell(i, 4).value,
            type=sheet_1.cell(i, 5).value,
            owner=sheet_1.cell(i, 6).value,
            topics=sheet_1.cell(i, 7).value,
            authors=sheet_1.cell(i, 8).value,
        )
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02
    style.alignment = al
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/软件著作权')
    path = os.path.join(path, '已导入软件著作权信息.xls')

    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '著作权名称', style)
        sheet.write(0, 1, '证书号', style)
        sheet.write(0, 2, '开发完成时间', style)
        sheet.write(0, 3, '获得时间', style)
        sheet.write(0, 4, '登记号', style)
        sheet.write(0, 5, '著作权类型', style)
        sheet.write(0, 6, '著作权人', style)
        sheet.write(0, 7, '关联课题', style)
        sheet.write(0, 8, '作者信息', style)
        work_book.save(path)
    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    table = work_book.sheets()[0]
    nrows = table.nrows
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    third_col = sheet.col(2)
    fifth_col = sheet.col(4)
    sixth_col = sheet.col(5)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    nineth_col = sheet.col(8)
    fifteenth_col = sheet.col(14)
    sixteenth_col = sheet.col(15)
    first_col.width = 256 * 60
    sec_col.width = 256 * 25
    third_col.width = 256 * 25
    fifth_col.width = 256 * 25
    sixth_col.width = 256 * 40
    seventh_col.width = 256 * 25
    eightth_col.width = 256 * 25
    nineth_col.width = 256 * 40
    fifteenth_col.width = 256 * 25
    sixteenth_col.width = 256 * 25
    for i in range(1, nrows_ori):
        sheet.write(nrows, 0, sheet_1.cell(i, 0).value)
        sheet.write(nrows, 1, sheet_1.cell(i, 1).value)
        sheet.write(nrows, 2, sheet_1.cell(i, 2).value)
        sheet.write(nrows, 3, sheet_1.cell(i, 3).value)
        sheet.write(nrows, 4, sheet_1.cell(i, 4).value)
        sheet.write(nrows, 5, sheet_1.cell(i, 5).value)
        sheet.write(nrows, 6, sheet_1.cell(i, 6).value)
        sheet.write(nrows, 7, sheet_1.cell(i, 7).value)
        sheet.write(nrows, 8, sheet_1.cell(i, 8).value)
        nrows += 1
    excel.save(path)


# 保存获奖信息
def save_AwardMessage(path, scholarId):
    work_book = xlrd.open_workbook(path)
    sheet_1 = work_book.sheet_by_index(0)
    table = work_book.sheets()[0]
    nrows_ori = table.nrows
    for i in range(1, nrows_ori):
        AwardMessage.objects.create(
            scholarId=scholarId,
            title=sheet_1.cell(i, 0).value,
            rank=sheet_1.cell(i, 1).value,
            level=sheet_1.cell(i, 2).value,
            org=sheet_1.cell(i, 3).value,
            getTime=int(sheet_1.cell(i, 4).value),
            topics=sheet_1.cell(i, 5).value,
            authors=sheet_1.cell(i, 6).value,
        )
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02
    style.alignment = al
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/获奖')
    path = os.path.join(path, '已导入获奖信息.xls')

    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '荣誉名称', style)
        sheet.write(0, 1, '获奖名次', style)
        sheet.write(0, 2, '级别', style)
        sheet.write(0, 3, '授予单位', style)
        sheet.write(0, 4, '获奖时间', style)
        sheet.write(0, 5, '关联课题', style)
        sheet.write(0, 6, '作者信息', style)
        work_book.save(path)
    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    table = work_book.sheets()[0]
    nrows = table.nrows
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    third_col = sheet.col(2)
    fourth_col = sheet.col(3)
    fifth_col = sheet.col(4)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    fifteenth_col = sheet.col(14)
    sixteenth_col = sheet.col(15)
    first_col.width = 256 * 40
    sec_col.width = 256 * 25
    third_col.width = 256 * 40
    fourth_col.width = 256 * 40
    fifth_col.width = 256 * 10
    seventh_col.width = 256 * 25
    eightth_col.width = 256 * 25
    fifteenth_col.width = 256 * 25
    sixteenth_col.width = 256 * 25
    for i in range(1, nrows_ori):
        sheet.write(nrows, 0, sheet_1.cell(i, 0).value)
        sheet.write(nrows, 1, sheet_1.cell(i, 1).value)
        sheet.write(nrows, 2, sheet_1.cell(i, 2).value)
        sheet.write(nrows, 3, sheet_1.cell(i, 3).value)
        sheet.write(nrows, 4, sheet_1.cell(i, 4).value)
        sheet.write(nrows, 5, sheet_1.cell(i, 5).value)
        sheet.write(nrows, 6, sheet_1.cell(i, 6).value)
        nrows += 1
    excel.save(path)


# 保存学生获奖信息
def save_StudentAwardMessage(path, scholarId):
    work_book = xlrd.open_workbook(path)
    sheet_1 = work_book.sheet_by_index(0)
    table = work_book.sheets()[0]
    nrows_ori = table.nrows
    for i in range(1, nrows_ori):
        StudentAwardMessage.objects.create(
            scholarId=scholarId,
            title=sheet_1.cell(i, 0).value,
            student=sheet_1.cell(i, 1).value,
            rank=sheet_1.cell(i, 2).value,
            level=sheet_1.cell(i, 3).value,
            org=sheet_1.cell(i, 4).value,
            getTime=int(sheet_1.cell(i, 5).value),
            topics=sheet_1.cell(i, 6).value,
            authors=sheet_1.cell(i, 7).value,
        )
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02
    style.alignment = al
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/学生获奖')
    path = os.path.join(path, '已导入学生获奖信息.xls')

    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '比赛名称', style)
        sheet.write(0, 1, '获奖学生', style)
        sheet.write(0, 2, '获奖名次', style)
        sheet.write(0, 3, '级别', style)
        sheet.write(0, 4, '授予单位', style)
        sheet.write(0, 5, '获奖时间', style)
        sheet.write(0, 6, '关联课题', style)
        sheet.write(0, 7, '作者信息', style)
        work_book.save(path)
    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    table = work_book.sheets()[0]
    nrows = table.nrows
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    fourth_col = sheet.col(3)
    fifth_col = sheet.col(4)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    fifteenth_col = sheet.col(14)
    sixteenth_col = sheet.col(15)
    first_col.width = 256 * 40
    sec_col.width = 256 * 40
    fourth_col.width = 256 * 25
    fifth_col.width = 256 * 40
    seventh_col.width = 256 * 25
    eightth_col.width = 256 * 25
    fifteenth_col.width = 256 * 25
    sixteenth_col.width = 256 * 25
    for i in range(1, nrows_ori):
        sheet.write(nrows, 0, sheet_1.cell(i, 0).value)
        sheet.write(nrows, 1, sheet_1.cell(i, 1).value)
        sheet.write(nrows, 2, sheet_1.cell(i, 2).value)
        sheet.write(nrows, 3, sheet_1.cell(i, 3).value)
        sheet.write(nrows, 4, sheet_1.cell(i, 4).value)
        sheet.write(nrows, 5, sheet_1.cell(i, 5).value)
        sheet.write(nrows, 6, sheet_1.cell(i, 6).value)
        sheet.write(nrows, 7, sheet_1.cell(i, 7).value)
        nrows += 1
    excel.save(path)


# 获取文件
def get_files(request):
    project_name = request.POST.get("project_name")
    scholarId = request.POST.get("scholarId")
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId)
    if not os.path.exists(path):
        os.makedirs(os.path.join(path, '论文'))
        os.makedirs(os.path.join(path, '专利'))
        os.makedirs(os.path.join(path, '项目'))
        os.makedirs(os.path.join(path, '学生获奖'))
        os.makedirs(os.path.join(path, '获奖'))
        os.makedirs(os.path.join(path, '软件著作权'))
        shutil.copy(os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/论文/论文模板.xlsx'),
                    os.path.join(path, '论文/论文模板.xls'))
        shutil.copy(os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/专利/专利模版.xlsx'),
                    os.path.join(path, '专利/专利模板.xls'))
        shutil.copy(os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/项目/横向项目模版.xlsx'),
                    os.path.join(path, '项目/横向项目模板.xls'))
        shutil.copy(os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/项目/纵向项目模版.xlsx'),
                    os.path.join(path, '项目/纵向项目模板.xls'))
        shutil.copy(os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/获奖/获奖模版.xlsx'),
                    os.path.join(path, '获奖/获奖模板.xls'))
        shutil.copy(os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/软件著作权/软件著作权模版.xlsx'),
                    os.path.join(path, '软件著作权/软件著作权.xls'))
        shutil.copy(os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/学生获奖/学生获奖模版.xlsx'),
                    os.path.join(path, '学生获奖/学生获奖模板.xls'))
    files = []
    if project_name != '项目':
        file = {
            'name': project_name + '模板.xlsx',
        }
        files.append(file)
    else:
        file1 = {
            'name': '横向项目模板.xlsx',
        }
        file2 = {
            'name': '纵向项目模板.xlsx',
        }
        files.append(file1)
        files.append(file2)
    for name in os.listdir(os.path.join(path, project_name)):
        file = {
            'name': name,
        }
        files.append(file)
    return JsonResponse({'files': files,
                         'data': model_to_dict(Imported_persons.objects.filter(scholarId=scholarId).first()),
                         "state": 1, "message": "返回文件成功"}, json_dumps_params={"ensure_ascii": False})


# 根据文件名访问文件
def get_FileByName(request):
    scholarId = request.GET.get("scholarId")
    project_name = request.GET.get("project_name")
    name = request.GET.get('name')
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if name in ['论文模板.xlsx', '专利模板.xlsx', '横向项目模板.xlsx', '纵向项目模板.xlsx', '软件著作权模板.xlsx',
                '获奖模板.xlsx', '学生获奖模板.xlsx']:
        path = BASE_DIR + '/MyApp/dist/files/教师个人端/成果管理模版/' + project_name
        path = os.path.join(path, name)
    else:
        path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/',
                            scholarId, project_name, name)
    with open(path, 'rb') as f:
        response = HttpResponse(content_type="application/vnd.ms-excel")
        response.write(f.read())
        response['Content-Disposition'] = 'attachment;filename={0}'.format(quote(name))
        return response


def add_Paper_management(request):
    paperIds = json.loads(request.POST.get('paperIds'))
    scholarId = request.POST.get('scholarId')
    for paperId in paperIds:
        if len(PaperManagement.objects.filter(paperId=paperId, scholarId=scholarId)) == 0:
            PaperManagement.objects.create(paperId=paperId, scholarId=scholarId)
            getPaperMessagesByPaperId(paperId, scholarId)
    paperCount = len(PaperManagement.objects.filter(scholarId=scholarId))
    if len(AchievementCount.objects.filter(scholarId=scholarId)) != 0:
        person = AchievementCount.objects.filter(scholarId=scholarId)
        person.update(paperCount=paperCount)
    else:
        AchievementCount.objects.create(scholarId=scholarId, paperCount=paperCount)
    return JsonResponse({"state": 1, "message": "数据已添加至论文相关表"}, json_dumps_params={"ensure_ascii": False})


def getPaperMessagesByPaperId(paperId, scholarId):
    url = 'https://zhitulist.com/academic/api/v1/papers/' + str(paperId)
    req = requests.get(url)
    data = json.loads(req.text)['data']
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    style.alignment = al
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/论文')
    path = os.path.join(path, '已导入论文信息.xls')
    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '论文名', style)
        sheet.write(0, 1, '摘要', style)
        sheet.write(0, 2, '发表年份', style)
        sheet.write(0, 3, '论文类型', style)
        sheet.write(0, 4, 'DOI号', style)
        sheet.write(0, 5, '语言', style)
        sheet.write(0, 6, '刊物名称', style)
        sheet.write(0, 7, '论文领域', style)
        sheet.write(0, 8, '议题', style)
        sheet.write(0, 9, '卷号', style)
        sheet.write(0, 10, '起始页码', style)
        sheet.write(0, 11, '结束页码', style)
        sheet.write(0, 12, '出版物', style)
        sheet.write(0, 13, '引用次数', style)
        sheet.write(0, 14, '关键词', style)
        sheet.write(0, 15, '作者信息', style)
        sheet.write(0, 16, '是否为sci', style)
        sheet.write(0, 17, '是否为ei', style)
        work_book.save(path)

    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    fifth_col = sheet.col(4)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    nineth_col = sheet.col(8)
    fifteenth_col = sheet.col(14)
    sixteenth_col = sheet.col(15)
    first_col.width = 256 * 80
    sec_col.width = 256 * 80
    fifth_col.width = 256 * 25
    seventh_col.width = 256 * 80
    eightth_col.width = 256 * 80
    nineth_col.width = 256 * 25
    fifteenth_col.width = 256 * 25
    sixteenth_col.width = 256 * 25

    table = work_book.sheets()[0]
    nrows = table.nrows
    sheet.write(nrows, 0, data['title'])
    sheet.write(nrows, 1, data['abst'])
    sheet.write(nrows, 2, data['year'])
    sheet.write(nrows, 3, data['docType'])
    sheet.write(nrows, 4, data['doi'])
    sheet.write(nrows, 5, data['lang'])
    sheet.write(nrows, 6, data['venue'])
    sheet.write(nrows, 7, data['keywords'])
    sheet.write(nrows, 8, data['issue'])
    sheet.write(nrows, 9, data['volume'])
    sheet.write(nrows, 10, data['pageStart'])
    sheet.write(nrows, 11, data['pageEnd'])
    sheet.write(nrows, 12, data['publisher'])
    sheet.write(nrows, 13, data['citationNum'])
    sheet.write(nrows, 14, data['keywords'])
    authors = ''
    for person in data['scholars']:
        authors += person['name'] + ' '
    sheet.write(nrows, 15, authors)
    if data['sciPaper']:
        sheet.write(nrows, 16, '是')
    else:
        sheet.write(nrows, 16, '否')
    if data['eiPaper']:
        sheet.write(nrows, 17, '是')
    else:
        sheet.write(nrows, 17, '否')
    excel.save(path)


def getPatentMessagesByPatentId(patentId, scholarId):
    url = 'https://zhitulist.com/academic/api/v1/patents/' + str(patentId)
    req = requests.get(url)
    data = json.loads(req.text)['data']
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    style.alignment = al
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/专利/已导入专利信息.xls')
    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '专利名称', style)
        sheet.write(0, 1, '专利类型', style)
        sheet.write(0, 2, '专利状态', style)
        sheet.write(0, 3, '专利编号', style)
        sheet.write(0, 4, '专利权人', style)
        sheet.write(0, 5, '授权公告日', style)
        sheet.write(0, 6, '申请编号', style)
        sheet.write(0, 7, '专利申请日', style)
        work_book.save(path)
    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    table = work_book.sheets()[0]
    nrows = table.nrows
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    third_col = sheet.col(2)
    fourth_col = sheet.col(3)
    fifth_col = sheet.col(4)
    sixth_col = sheet.col(5)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    first_col.width = 256 * 80
    sec_col.width = 256 * 25
    third_col.width = 256 * 25
    fourth_col.width = 256 * 25
    fifth_col.width = 256 * 25
    sixth_col.width = 256 * 25
    seventh_col.width = 256 * 25
    eightth_col.width = 256 * 40

    sheet.write(nrows, 0, data['title'])
    sheet.write(nrows, 1, data['patentType'])
    sheet.write(nrows, 2, data['legalStatus'])
    sheet.write(nrows, 3, data['authorizationNum'])
    sheet.write(nrows, 4, data['inventorName'])
    sheet.write(nrows, 5, data['priorityDate'])
    sheet.write(nrows, 6, data['applicationNum'])
    sheet.write(nrows, 7, data['applicationDate'])
    excel.save(path)


def getProjectMessagesByProjectId(projectId, scholarId):
    url = 'https://zhitulist.com/academic/api/v1/projects/' + str(projectId)
    req = requests.get(url)
    data = json.loads(req.text)['data']
    style = xlwt.easyxf('align: wrap on')
    al = xlwt.Alignment()
    al.horz = 0x02  # 设置水平居中
    style.alignment = al

    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    path = os.path.join(BASE_DIR, 'MyApp/dist/files/教师个人端/成果管理模版/' + scholarId + '/项目/已导入项目信息.xls')
    if not os.path.exists(path):
        work_book = xlwt.Workbook()
        sheet = work_book.add_sheet('sheet')
        sheet.write(0, 0, '课题名称', style)
        sheet.write(0, 1, '合同来源', style)
        sheet.write(0, 2, '合同编号', style)
        sheet.write(0, 3, '基金名称', style)
        sheet.write(0, 4, '授予单位', style)
        sheet.write(0, 5, '课题级别（一级）', style)
        sheet.write(0, 6, '课题级别（二级）', style)
        sheet.write(0, 7, '到账经费（万）', style)
        sheet.write(0, 8, '开始时间', style)
        sheet.write(0, 9, '截止时间', style)
        sheet.write(0, 10, '项目状态', style)
        sheet.write(0, 11, '作者信息', style)
        work_book.save(path)
    work_book = xlrd.open_workbook(path)
    excel = copy(wb=work_book)
    sheet = excel.get_sheet(0)
    table = work_book.sheets()[0]
    nrows = table.nrows
    first_col = sheet.col(0)
    sec_col = sheet.col(1)
    third_col = sheet.col(2)
    fourth_col = sheet.col(3)
    fifth_col = sheet.col(4)
    sixth_col = sheet.col(5)
    seventh_col = sheet.col(6)
    eightth_col = sheet.col(7)
    first_col.width = 256 * 40
    sec_col.width = 256 * 25
    third_col.width = 256 * 25
    fourth_col.width = 256 * 25
    fifth_col.width = 256 * 25
    sixth_col.width = 256 * 25
    seventh_col.width = 256 * 25
    eightth_col.width = 256 * 25

    sheet.write(nrows, 0, data['title'])
    sheet.write(nrows, 1, '无')
    sheet.write(nrows, 2, '无')
    sheet.write(nrows, 3, '无')
    sheet.write(nrows, 4, data['org'])
    sheet.write(nrows, 5, data['typeFirst'])
    sheet.write(nrows, 6, data['typeSecondary'])
    sheet.write(nrows, 7, data['fund'])
    sheet.write(nrows, 8, data['startYear'])
    sheet.write(nrows, 9, data['endYear'])
    sheet.write(nrows, 10, '无')
    sheet.write(nrows, 11, data['leader'])
    excel.save(path)


def add_Project_management(request):
    projectIds = json.loads(request.POST.get('projectIds'))
    scholarId = request.POST.get('scholarId')
    for projectId in projectIds:
        if len(ProjectManagement.objects.filter(projectId=projectId, scholarId=scholarId)) == 0:
            ProjectManagement.objects.create(projectId=projectId, scholarId=scholarId)
            getProjectMessagesByProjectId(projectId, scholarId)
    projectCount = len(ProjectManagement.objects.filter(scholarId=scholarId))
    if len(AchievementCount.objects.filter(scholarId=scholarId)) != 0:
        person = AchievementCount.objects.filter(scholarId=scholarId)
        person.update(projectCount=projectCount)
    else:
        AchievementCount.objects.create(scholarId=scholarId, projectCount=projectCount)
    return JsonResponse({"state": 1, "message": "数据已添加至项目相关表"}, json_dumps_params={"ensure_ascii": False})


def add_Patent_management(request):
    patentIds = json.loads(request.POST.get('patentIds'))
    scholarId = request.POST.get('scholarId')
    for patentId in patentIds:
        if len(PatentManagement.objects.filter(patentId=patentId, scholarId=scholarId)) == 0:
            PatentManagement.objects.create(patentId=patentId, scholarId=scholarId)
            getPatentMessagesByPatentId(patentId, scholarId)
    patentCount = len(PatentManagement.objects.filter(scholarId=scholarId))
    if len(AchievementCount.objects.filter(scholarId=scholarId)) != 0:
        person = AchievementCount.objects.filter(scholarId=scholarId)
        person.update(patentCount=patentCount)
    else:
        AchievementCount.objects.create(scholarId=scholarId, patentCount=patentCount)
    return JsonResponse({"state": 1, "message": "数据已添加至专利相关表"}, json_dumps_params={"ensure_ascii": False})


def get_Paper(request):
    scholarId = request.POST.get('scholarId')
    list = []
    for paper in PaperManagement.objects.filter(scholarId=scholarId):
        list.append(paper.paperId)
    return JsonResponse({'data': list, "state": 1,
                         "message": "数据已经成功返回"}, json_dumps_params={"ensure_ascii": False})


def get_Patent(request):
    scholarId = request.POST.get('scholarId')
    list = []
    for patent in PatentManagement.objects.filter(scholarId=scholarId):
        list.append(patent.patentId)
    return JsonResponse({'data': list, "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_Project(request):
    scholarId = request.POST.get('scholarId')
    list = []
    for project in ProjectManagement.objects.filter(scholarId=scholarId):
        list.append(project.projectId)
    return JsonResponse({'data': list, "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_All_Papers(request):
    scholarId = request.POST.get('scholarId')
    total_papers = int(request.POST.get('total_papers'))
    page = 0
    papers = []
    while (total_papers > 0):
        url = 'https://zhitulist.com/academic/api/v1/scholars/' + scholarId + '/papers?page=' + str(page) + '&num=30'
        req = requests.get(url)
        page += 1
        total_papers -= 30
        papers.extend(json.loads(req.text)['data']['content'])
    papers2 = []
    for papermessage in PaperMessage.objects.filter(scholarId=scholarId):
        papers2.append(model_to_dict(papermessage))
    return JsonResponse({'data': papers, 'data2': papers2, "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_All_Patents(request):
    scholarId = request.POST.get('scholarId')
    total_patents = int(request.POST.get('total_patents'))
    page = 0
    patents = []
    while (total_patents > 0):
        url = 'https://zhitulist.com/academic/api/v1/scholars/' + scholarId + '/patents?page=' + str(page) + '&num=30'
        req = requests.get(url)
        page += 1
        total_patents -= 30
        patents.extend(json.loads(req.text)['data']['content'])
    patents2 = []
    for patentmessage in PatentMessage.objects.filter(scholarId=scholarId):
        patents2.append(model_to_dict(patentmessage))
    return JsonResponse({'data': patents, 'data2': patents2, "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_All_Projects(request):
    scholarId = request.POST.get('scholarId')
    total_projects = int(request.POST.get('total_projects'))
    page = 0
    projects = []
    while (total_projects > 0):
        url = 'https://zhitulist.com/academic/api/v1/scholars/' + scholarId + '/projects?page=' + str(page) + '&num=30'
        req = requests.get(url)
        page += 1
        total_projects -= 30
        projects.extend(json.loads(req.text)['data']['content'])
    projects2 = []
    for projectmessage in ProjectVerticalMessage.objects.filter(scholarId=scholarId):
        projects2.append(model_to_dict(projectmessage))
    for projectmessage in ProjectTransverseMessage.objects.filter(scholarId=scholarId):
        projects2.append(model_to_dict(projectmessage))
    return JsonResponse({'data': projects,
                         'data2': projects2,
                         "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_All_SoftwareCopyright(request):
    scholarId = request.POST.get('scholarId')
    SoftwareCopyrights = []
    for softwarecopyright in SoftwareCopyrightMessage.objects.filter(scholarId=scholarId):
        SoftwareCopyrights.append(model_to_dict(softwarecopyright))
    return JsonResponse({'data': SoftwareCopyrights,
                         "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_All_Award(request):
    scholarId = request.POST.get('scholarId')
    Awards = []
    for award in AwardMessage.objects.filter(scholarId=scholarId):
        Awards.append(model_to_dict(award))
    return JsonResponse({'data': Awards,
                         "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_All_StudentAward(request):
    scholarId = request.POST.get('scholarId')
    StudentAwards = []
    for studentaward in StudentAwardMessage.objects.filter(scholarId=scholarId):
        StudentAwards.append(model_to_dict(studentaward))
    return JsonResponse({'data': StudentAwards,
                         "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def change_Person_State(request):
    scholarId = request.POST.get('scholarId')
    person = Admin_messages.objects.filter(scholarId=scholarId)
    person.update(state=1)
    return JsonResponse({"message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


# 根据论文id返回论文详细信息
def get_All_Papers_By_Ids(request):
    paperIds = json.loads(request.POST.get("paperIds"))
    papers = []
    for paperId in paperIds:
        url = 'https://zhitulist.com/academic/api/v1/papers/' + paperId
        req = requests.get(url)
        papers.append(json.loads(req.text)['data'])
    return JsonResponse({'data': papers, "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def get_AchievementReportDetailByScholarId(request):
    scholarId = request.POST.get('scholarId')
    begin_year = request.POST.get('begin_year')
    end_year = request.POST.get('end_year')
    managerId = request.POST.get('managerId')
    report_id = Achievement_report.objects.filter(begin_year=begin_year, end_year=end_year, managerId=managerId).first()
    data = model_to_dict(Achievement_report_detail.objects.filter(
        report_id=report_id, scholarId=scholarId, managerId=managerId).first())
    return JsonResponse({'data': data, "state": 1,
                         "message": "数据已经成功返回"},
                        json_dumps_params={"ensure_ascii": False})


def change_AchievementReportDetailStateByScholarId(request):
    scholarId = request.POST.get('scholarId')
    begin_year = request.POST.get('begin_year')
    end_year = request.POST.get('end_year')
    managerId = request.POST.get('managerId')
    report_id = Achievement_report.objects.filter(
        begin_year=begin_year, end_year=end_year, managerId=managerId).first()
    person = Achievement_report_detail.objects.filter(
        report_id=report_id, scholarId=scholarId, managerId=managerId)
    person.update(state=1)
    return JsonResponse({'data': model_to_dict(person.first()), "state": 1,
                         "message": "相关数据信息已确认"},
                        json_dumps_params={"ensure_ascii": False})


def get_ScholarsByYear(request):
    data = json.loads(request.POST.get('data'))
    managerId = request.POST.get("managerId")
    begin_year = data['begin_year']
    end_year = data['end_year']
    report_id = Achievement_report.objects.filter(begin_year=begin_year, end_year=end_year, managerId=managerId).first()
    persons = Achievement_report_detail.objects.filter(report_id=report_id, managerId=managerId)
    list = []
    for person in persons:
        list.append(model_to_dict(person))
    return JsonResponse({'data': list, "state": 1,
                         "message": "相关数据信息已确认"},
                        json_dumps_params={"ensure_ascii": False})


# 验证是否为使用者
def judgeManageId(request):
    managerId = request.POST.get('managerId')
    if len(orgManager.objects.filter(managerId=managerId)) != 0:
        orgmanager = orgManager.objects.filter(managerId=managerId).first()
        return JsonResponse({"data": model_to_dict(orgmanager),
                             "state": 1,
                             "message": "该使用者存在"},
                            json_dumps_params={"ensure_ascii": False})
    else:
        return JsonResponse({"state": 0,
                             "message": "该使用者并不存在，要想使用该系统请联系17373255@buaa.edu.cn"},
                            json_dumps_params={"ensure_ascii": False})


# 修改管理人员状态
def changeManagerState(request):
    state = int(request.POST.get('state'))
    managerId = request.POST.get("managerId")
    orgManager.objects.filter(managerId=managerId).update(state=state)
    return JsonResponse({"message": "管理人员状态修改完成"},
                        json_dumps_params={"ensure_ascii": False})

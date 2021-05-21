from django.db import models


# Create your models here.

# 待导入成员信息表（工号、姓名、部门、邮箱、职称、学校）
class Wait_persons(models.Model):
    id1 = models.CharField(max_length=20,null=False)
    name = models.CharField(max_length=20, null=False)
    department = models.CharField(max_length=20, null=False)
    email = models.CharField(max_length=40, null=False)
    title = models.CharField(max_length=20, null=False)
    orgName = models.CharField(max_length=30, null=False)
    managerId = models.CharField(max_length=20)


# 已导入成员信息表
class Imported_persons(models.Model):
    id1 = models.CharField(max_length=20,null=False)
    name = models.CharField(max_length=20, null=False)
    department = models.CharField(max_length=20, null=False)
    email = models.CharField(max_length=40, null=False)
    title = models.CharField(max_length=20, null=False)
    orgName = models.CharField(max_length=30, null=False)
    avg = models.CharField(max_length=100, null=False)
    scholarId = models.CharField(max_length=100, null=False)
    paperCount = models.IntegerField(null=False)
    projectCount = models.IntegerField(null=False)
    patentCount = models.IntegerField(null=False)
    awardCount = models.IntegerField(null=False, default=0)
    student_awardCount = models.IntegerField(null=False, default=0)
    workCount = models.IntegerField(null=False, default=0)
    copyrightCount = models.IntegerField(null=False, default=0)
    managerId = models.CharField(max_length=20,default=0)


# 年度成果报告
class Achievement_report(models.Model):
    report_id = models.AutoField(primary_key=True)
    begin_year = models.IntegerField(null=False)
    end_year = models.IntegerField(null=False)
    managerId = models.CharField(max_length=20)


# 这里建造一个成果报告详情表，外键对应年度成果报告，这样就可以有效解决相关问题了
class Achievement_report_detail(models.Model):
    report_detail_id = models.AutoField(primary_key=True)
    report_id = models.ForeignKey(Achievement_report, on_delete=models.CASCADE)
    state = models.IntegerField(null=False, default=0)
    id1 = models.CharField(max_length=20, null=False, default=0)
    name = models.CharField(max_length=20, null=False, default=0)
    email = models.CharField(max_length=40, null=False, default=0)
    # # 用的时候记得把scholarId转成数字
    scholarId = models.CharField(max_length=100, null=False, default=0)
    paperCount = models.IntegerField(null=False, default=0)
    paperSciCount = models.IntegerField(null=False, default=0)
    paperEiCount = models.IntegerField(null=False, default=0)
    paperOtherCount = models.IntegerField(null=False, default=0)
    projectCount = models.IntegerField(null=False, default=0)
    projectNationCount = models.IntegerField(null=False, default=0)
    projectProvinceCount = models.IntegerField(null=False, default=0)
    projectOtherCount = models.IntegerField(null=False, default=0)
    patentCount = models.IntegerField(null=False, default=0)
    awardCount = models.IntegerField(null=False, default=0)
    awardNationCount = models.IntegerField(null=False, default=0)
    awardProvinceCount = models.IntegerField(null=False, default=0)
    awardOtherCount = models.IntegerField(null=False, default=0)
    student_awardCount = models.IntegerField(null=False, default=0)
    student_awardNationCount = models.IntegerField(null=False, default=0)
    student_awardProvinceCount = models.IntegerField(null=False, default=0)
    student_awardOtherCount = models.IntegerField(null=False, default=0)
    workCount = models.IntegerField(null=False, default=0)
    software_copyrightCount = models.IntegerField(null=False, default=0)
    managerId = models.CharField(max_length=20,default=0)

# 学者人员相关账号(邮箱作为账号)
class Admin_messages(models.Model):
    email = models.CharField(max_length=100, primary_key=True)
    password = models.CharField(max_length=100, null=False)
    scholarId = models.CharField(max_length=20, null=False)
    state = models.IntegerField(null=False, default=0)


# 论文管理
class PaperManagement(models.Model):
    paperId = models.CharField(max_length=20)
    scholarId = models.CharField(max_length=20)


# 专利管理
class PatentManagement(models.Model):
    patentId = models.CharField(max_length=20)
    scholarId = models.CharField(max_length=20)


# 项目管理
class ProjectManagement(models.Model):
    projectId = models.CharField(max_length=20)
    scholarId = models.CharField(max_length=20)


# 已确认成果数
class AchievementCount(models.Model):
    scholarId = models.CharField(max_length=20, primary_key=True)
    paperCount = models.IntegerField(null=False, default=0)
    projectCount = models.IntegerField(null=False, default=0)
    patentCount = models.IntegerField(null=False, default=0)
    workCount = models.IntegerField(null=False, default=0)
    awardCount = models.IntegerField(null=False, default=0)
    studentAwardCount = models.IntegerField(null=False, default=0)
    copyrightCount = models.IntegerField(null=False, default=0)


# 论文信息
class PaperMessage(models.Model):
    scholarId = models.CharField(max_length=20)
    paperId = models.AutoField(primary_key=True)
    title = models.TextField()
    abst = models.TextField()
    year = models.IntegerField()
    docType = models.TextField()
    doi = models.TextField()
    lang = models.TextField()
    venue = models.TextField()
    filed = models.TextField(default='无')
    issue = models.IntegerField()
    volume = models.IntegerField()
    pageStart = models.IntegerField()
    pageEnd = models.IntegerField()
    publisher = models.TextField()
    citationNum = models.IntegerField()
    keywords = models.TextField()
    authors = models.TextField()
    sciPaper = models.BooleanField(default=False)
    eiPaper = models.BooleanField(default=False)


# 专利信息
class PatentMessage(models.Model):
    scholarId = models.CharField(max_length=20)
    patentId = models.AutoField(primary_key=True)
    title = models.TextField()
    patentType = models.TextField()
    legalStatus = models.TextField()
    authorizationNum = models.TextField()
    inventorName = models.TextField()
    priorityDate = models.TextField()
    applicationNum = models.TextField()
    applicationDate = models.TextField()

# 横向项目信息
class ProjectTransverseMessage(models.Model):
    scholarId = models.CharField(max_length=20)
    projectId = models.AutoField(primary_key=True)
    title = models.TextField()
    source = models.TextField()
    contractId = models.TextField()
    fund = models.FloatField()
    startYear = models.IntegerField()
    endYear = models.IntegerField()
    projectStatus = models.CharField(max_length=20)
    authors = models.TextField()
    typeFirst = models.TextField(default='其他项目')
    typeSecondary = models.TextField(default='其他项目')

# 纵向项目信息
class ProjectVerticalMessage(models.Model):
    scholarId = models.CharField(max_length=20)
    projectId = models.AutoField(primary_key=True)
    title = models.TextField()
    typeTertiary = models.TextField()
    org = models.TextField()
    typeFirst = models.TextField()
    typeSecondary = models.TextField()
    fund = models.FloatField()
    startYear = models.IntegerField()
    endYear = models.IntegerField()
    projectStatus = models.CharField(max_length=20)
    authors = models.TextField()


# 软件著作权
class SoftwareCopyrightMessage(models.Model):
    scholarId = models.CharField(max_length=20)
    projectId = models.AutoField(primary_key=True)
    title = models.TextField()
    certificateId = models.TextField()
    endTime = models.IntegerField()
    getTime = models.IntegerField()
    registrationNum = models.TextField()
    type = models.TextField()
    owner = models.TextField()
    topics = models.TextField()
    authors = models.TextField()


# 获奖
class AwardMessage(models.Model):
    scholarId = models.CharField(max_length=20)
    projectId = models.AutoField(primary_key=True)
    title = models.TextField()
    rank = models.TextField()
    level = models.TextField()
    org = models.TextField()
    getTime = models.IntegerField()
    topics = models.TextField()
    authors = models.TextField()


# 学生获奖
class StudentAwardMessage(models.Model):
    scholarId = models.CharField(max_length=20)
    projectId = models.AutoField(primary_key=True)
    title = models.TextField()
    student = models.TextField()
    rank = models.TextField()
    level = models.TextField()
    org = models.TextField()
    getTime = models.IntegerField()
    topics = models.TextField()
    authors = models.TextField()


# 管理人员账号
class orgManager(models.Model):
    managerId = models.CharField(max_length=20,primary_key=True)
    orgName = models.CharField(max_length=40)
    department = models.CharField(max_length=40)
    manageName = models.CharField(max_length=40)
    state = models.IntegerField() #用于判断是否已经完成导入成员

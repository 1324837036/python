from django.urls import path
from MyApp import views

urlpatterns = [
    path("test/", views.test),		# 添加一个新的路由
    path("add_Wait_persons/", views.add_Wait_persons),
    path('get_Wait_persons/', views.get_Wait_persons),
    path('add_Imported_persons/',views.add_Imported_persons),
    path('get_Imported_persons/',views.get_Imported_persons),
    path('remove_Imported_persons/',views.remove_Imported_persons),
    path('change_Imported_persons/', views.change_Imported_persons),
    path('send_emails/', views.send_emails),
    path("search_Imported_persons/",views.search_Imported_persons),
    path('add_Achievement_report/',views.add_Achievement_report),
    path('add_Achievement_report_detail/', views.add_Achievement_report_detail),
    path('get_Achievement_report/',views.get_Achievement_report),
    path('get_Achievement_report_detail/',views.get_Achievement_report_detail),
    path('add_Admin_messages/',views.add_Admin_messages),
    path('send_email_Single/',views.send_email_Single),
    path('get_Excel/', views.get_Excel),
    path('get_Excel2/', views.get_Excel2),
    path('login/',views.login),
    path('get_messageByScholarId/',views.get_messageByScholarId),
    path('save_file/',views.save_file),
    path('add_Paper_management/',views.add_Paper_management),
    path('add_Project_management/',views.add_Project_management),
    path('add_Patent_management/',views.add_Patent_management),
    path('get_Paper/',views.get_Paper),
    path('get_Patent/',views.get_Patent),
    path('get_Project/',views.get_Project),
    path('get_All_Papers/',views.get_All_Papers),
    path('get_All_Patents/',views.get_All_Patents),
    path('get_All_Projects/',views.get_All_Projects),
    path('change_Person_State/',views.change_Person_State),
    path('get_files/',views.get_files),
    path('get_FileByName/',views.get_FileByName),
    path('get_All_Papers_By_Ids/',views.get_All_Papers_By_Ids),
    path('get_FolderNum/',views.get_FolderNum),
    path('get_All_SoftwareCopyright/',views.get_All_SoftwareCopyright),
    path('get_All_Award/',views.get_All_Award),
    path('get_All_StudentAward/',views.get_All_StudentAward),
    path('get_AchievementReportDetailByScholarId/',views.get_AchievementReportDetailByScholarId),
    path('change_AchievementReportDetailStateByScholarId/',views.change_AchievementReportDetailStateByScholarId),
    path('send_Achivement_email_Single/',views.send_Achivement_email_Single),
    path('send_Achivement_emails/',views.send_Achivement_emails),
    path('get_ScholarsByYear/',views.get_ScholarsByYear),
    path('judgeManageId/',views.judgeManageId),
    path('changeManagerState/',views.changeManagerState),
]


from jira import JIRA
import time
import xlrd,xlsxwriter
from xpinyin import Pinyin

#bug信息dict结构
# issue_dict= {
#     'project': {'id': 10002},
#     'summary': 'test标题',
#     'description': 'test描述',
#     'issuetype': {'id': '10004'},
#     'components':[{'name': 'test'}],
#     'priority': {'id': '4'},
#     'assignee':{'name': 'machaoran'}
# }
#将用户名转化为拼音
def change_to_pinyin(username):
    p = Pinyin()
    if p.get_pinyin(username,'') == 'zhangnuo':
        return 'zhangna'
    else:
        return p.get_pinyin(username,'')
#获取bug类型
def get_issuetype(issuetype):
    issuetypes = {'新需求':'10007','缺陷':'10004','数据修改':'10100','需求问题':'10101','无效问题':'10102'}
    return issuetypes[issuetype]
#创建bug
def jira_create_issue(issue_dict):
    jira = JIRA(server='**',basic_auth=('machaoran', '111111'))
    # new_issue = jira.create_issue(fields=issue_dict)
    
    if issue_dict['project'] ==  {'id': '10302'}:
        project_name = 'CPDJ'
    else:
        project_name = 'XSBUG'
    issues = jira.search_issues(jql_str='project = %s AND summary ~"%s"'  %(project_name,issue_dict['summary']), maxResults=10,fields='comment')
    if issues.total == 0:
        # new_issue = jira.create_issue(fields=issue_dict)
        return("1")
    else:
        return("0")#查询到匹配的，不在创建
#获取项目编号
def get_project (project_name):
    projects = {'产品对接群':'10302','线上bug':'10005'}
    return projects[project_name]
#读取excel
def open_excel(file_path):
    #打开excel
    bk = xlrd.open_workbook(file_path)
    #打开sheet页 
    sh= bk.sheet_by_name('工作表2')
    #获取总行数 
    row_num = sh.nrows
    #获取总列数
    ncols = sh.ncols

    for  i  in range(1,row_num):
        #读取excel 将对应单元格的值绑定到eict中
        #绑定成功后  调jira_create_issue()创建bug
        row_data = {}
        row_data['project'] =  {'id': get_project(sh.cell_value(0,10))}
        if row_data['project'] == {'id':'10005'}:
            row_data['summary'] = ('%s-%s' %(str(round(sh.cell_value(i,0))),sh.cell_value(i,1).replace('\n','')))
        if row_data['project'] == {'id':'10302'}:
            row_data['summary'] = sh.cell_value(i,1).replace('\n','')
        row_data['description'] = sh.cell_value(i,1)
        row_data['issuetype'] =  {'id':get_issuetype(sh.cell_value(i,3))}
        row_data['components'] = [{'name': '家长端'}]
        row_data['priority'] = {'id': '4'}
        row_data['assignee']={'name': change_to_pinyin(sh.cell_value(i,8))}
        
        # print(row_data)
        # print('第%s条bug，创建完成!' %(i))
        if jira_create_issue(row_data) == '1':
            print('第%s条bug，创建完成!' %(i))
        else:
            print('第%s条bug，已存在，创建失败!' %(i))



if __name__ == '__main__':
    file_path = '/Users/houquan/Desktop/值班问题记录/未命名文件夹/动因体育问题2群.xlsx'
    # file_path = '/Users/houquan/Desktop/值班问题记录/动因体育线上问题2018-09-14副本.xlsx'
    open_excel(file_path)

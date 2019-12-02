# *===================================*
# -*- coding: utf-8 -*-
# * Time : 2019/11/8 15:25
# * Author : zhangsf
# *===================================*
import datetime
import requests
import re
import json
import ssl
import xlwt
#关闭了verify之后，引入urllib3的disable_warnings
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
ssl._create_default_https_context = ssl._create_unverified_context
class JDEmail:
    def __init__(self):
        pass
#抽成一个方法将返回归一化之后的list
def outputToList(data1):
    RULE_1 = "【用例评审结果】"
    RULE_2 = "【提测】"
    RULE_3 = "【测试报告】"
    RULE_4 = "【上线自查结果】"
    RULE = [RULE_1, RULE_2, RULE_3, RULE_4]
    json_str = json.loads(data1.text)
    pattern="(?<='Subject': ').*?(?=',)"
    pattern2="(?<= 'Sender': {'Mailbox': {'Name': ').*?(?=',)"
    #正则匹配到所有的主题
    #findall需要传入的格式为str格式，需要str（）
    all_subject=re.findall(pattern, str(json_str))
    print("获取到的主题数："+str(len(all_subject)))

    #找到特定的subject和其下标
    for index, item in enumerate(all_subject):
        print(item)
        end = 0
        #如果邮件中存在【用例评审结果】获取发件人作为负责人
        if RULE_1 in item:
            # print(index,item)
            end=index+1
            a = all_subject[index]
            #防止最后一个
            if end<len(all_subject):
                b=all_subject[end]
                tmp=str(json_str).split("'Subject': '"+a)[1].split("'Subject': '"+b)[0]
            else:
                tmp = str(json_str).split("'Subject': '" + a)[1]
            subject_person = re.findall(pattern2, str(tmp))
            all_subject[index]=all_subject[index]+"----------"+str(subject_person[0])
    out_subject=[]
    value=[0,0,0,0]
    key=''
    print("------------------根据RULE中规则过滤出来的邮件------------------------")
    for subject in all_subject:
        #针对一个邮件主题进行四个rule的检测，最终生成key为邮件，
        # value为4个数字（0,0,0,0）代表四个规则一个也没有匹配上
        for i in range(0, len(RULE)):
            if RULE[i] in subject:
                # print(i,RULE[i],subject)
                print(subject)
                # 获取到当前邮件的项目名,命名不符合规范过滤掉
                if len(subject.split(RULE[i])[0]) > 0:
                    key = subject.split(RULE[i])[0] + subject.split(RULE[i])[1]
                    # 第几个规则匹配成功就将第几个的value置为1
                    value[i] = 1
        #能获取到项目名的开始存储
        if key!="":
            #必须使用不可变的，需要转为tuple类型
            value2=tuple(value)
            map1=[key,value2]
            #将每一个项目添加到最终需要添加到Excel表格中
            out_subject.append(map1)
        key = ''
        #重新归为初始值
        value=[0,0,0,0]
    return out_subject
#将获取到的list数据导入到Excel表格中
def ListToExcel(email_dict):
    print("开始构建Excel表格")
    myxls = xlwt.Workbook()
    sheet1 = myxls.add_sheet(u'sheet1', cell_overwrite_ok=True)
    #先初始化Excel表格的表头的内容
    sheet1.write(0, 0, "编号")
    sheet1.write(0, 1, "项目")
    sheet1.write(0, 2, "用例评审结果邮件")
    sheet1.write(0, 3, "提测邮件")
    sheet1.write(0, 4, "测试报告邮件")
    sheet1.write(0, 5, "上线自查邮件")
    sheet1.write(0, 6, "负责人")
    sheet1.write(0, 7, "流程规范打分(总分4分)")
    #将所有的统计好的邮件信息填入到Excel表格中
    count = 0
    # 遍历email_dict
    for key in email_dict:
        subject_person=''
        if key.__contains__("----------"):
            subject=key.split("----------")[0]
            subject_person=key.split("----------")[1]
        else:
            subject = key
        subject_value = list(email_dict[key])
        for i in range(0, len(subject_value)):
            if subject_value[i]>1:
                subject_value[i]=1
        sum_count=subject_value[0]+subject_value[1]+subject_value[2]+subject_value[3]
        sheet1.write(count + 1, 0, count + 1)
        sheet1.write(count + 1, 1, subject)
        sheet1.write(count + 1, 2, subject_value[0])
        sheet1.write(count + 1, 3, subject_value[1])
        sheet1.write(count + 1, 4, subject_value[2])
        sheet1.write(count + 1, 5, subject_value[3])
        #填充的为负责人
        sheet1.write(count + 1, 6, subject_person)
        #填充的为总的评分
        sheet1.write(count + 1, 7, sum_count)
        count = count + 1
    file_name=datetime.date.today()
    print("Excel表名："+str(file_name))
    myxls.save('E:\\python_file\\web_spider\\'+str(file_name)+'.xls')
    print("结果统计完成！！！")
#将两部分的list数据汇总key相同的部分合并一下，保证最终导入到Excel表中的key唯一
def MergeData(data):
    #todo
    #存在一个问题邮件汇总如果出现了两封邮件是属于重发的情况，就会统计两次，解决办法超过1的都算1
    for item in data:
        for tmp in data:
            item[0].strip()
            tmp[0].strip()
            if item[0].__contains__(tmp[0]) or tmp[0].__contains__(item[0]):
                #后面添加必须包含"----------"是因为在发邮件的时候命名中存在不规范的时候，覆盖掉这次，
                #下次真正需要统一命名规范的时候就会统计错
                if len(item[0]) >= len(tmp[0]) :
                    if item[0].__contains__("----------"):
                     tmp[0] = item[0]
                elif tmp[0].__contains__("----------"):
                    item[0] = tmp[0]
    print("***********************合并之后的数据**********************")
    for i in data:
        print(i)
    print("**********************************************************")
    empty_dict = dict()
    for d in data:
        a = list(d[1])
        #判断字典中是否存在key存在就更新value，否则添加k-v
        if(empty_dict.__contains__(d[0])):
            empty_dict[d[0]][0] = empty_dict[d[0]][0] + a[0]
            empty_dict[d[0]][1] = empty_dict[d[0]][1] + a[1]
            empty_dict[d[0]][2] = empty_dict[d[0]][2] + a[2]
            empty_dict[d[0]][3] = empty_dict[d[0]][3] + a[3]
        else:
            empty_dict[d[0]]=a
    return  empty_dict
if __name__ =='__main__':
    #第一部分数据
    url = "https://com/owa/sessiondata.ashx"
    header = {
        'Cookie':'X-BackEndCookie=S-1-5-21-1713849901-2797640346-4150151575-1009910=u56Lnp2ejJqBmszJzcrKzMjSyJnIztLLzMac0sfMz8bSz8mZmcmezsnNy8qegYHNz87G0s7N0s3Gq8/JxczOxc7N; __jdu=1571023082727881273581; pin=zhangvalue; unick=zhangvalue; _tp=cOR1hNRYc3ccyW4KIJwFQQ%3D%3D; _pst=zhangvalue; shshshfpa=b77ba96a-37d4-9aed-2808-8f342d4876af-1571026704; ClientId=A84FA33F1BDC45AD95097F44E4AA1CB9; X-OWA-JS-PSD=1; RoutingKeyCookie=v2:%2fV1xBTCDLP7eJXStxne5F48i2Y5XrC1%2fglQ5zDPnTSQ%3d:d89b6794-bf46-480f-bcf3-4fe09974a4b1@jd.com; unpl=V2_ZzNtbUdTQEZ2W05WexFZDGILRwhKXhMQc19DAHsRXw1mBEJYclRCFX0URlRnGFgUZwEZX0RcQRxFCEdkeB5fA2AFEFlBZxpFK0oYEDlNDEY1WnxZRldAFXEIQlF6KWwGZzMSXHJXRxN1CUVWehBfA2QFG1RCVEUXcQFGZEsebDVXBBtfRVZLJXQ4R2Q5TQAFZQoUXkEaQxFzCEdXeRhVBmEAFFRLV0ATdwxPVEsYbAY%3d; shshshfpb=hK9LDqjkpSdZgURqoxXh1BA%3D%3D; TrackID=1nDRau2MeyWhemIqFXhYVrMilvDeaW_L9GTWtInP3XkUb_pHlROI3gqDILHRvmKnGNWOChdplj7rribkHaGLfbJg5jqHxWzJJJgQFiXmsI1A; __jdv=122132179|ssa.jd.com|-|referral|-|1574240821489; PCSYCityID=CN_110000_110100_110112; shshshfp=ea884eee0c571bd8c126b4842541c636; areaId=1; ipLoc-djd=1-2809-51216-0; 3AB9D23F7A4B3C9B=JZTUZRJV76NL6RWD6IMD2ULQP7WOUEYQ6FF4TIGNXZ2H7USWHQM6CBLV6MHV44GBEQXLJXPDMUQINRMNQZLBHHFDWM; jd.erp.lang=zh_CN; lang=zh-CN; umail.jd.com=cb43089295b786f1781d9fc614eb7aa670b0dfb5; amail.jd.com=zhangshengfan; AppcacheVer=15.1.1415.2:zh-cnbase; erp1.jd.com=A55DD2272F705F1663148F245EAB067B35A94E4C8E566EBEAFE6637DECAA28E4FFB299A420E0BC6D9A7960B28FA8146E68C8056E41A64A858DA64EDE393D39E9DCD9C53A6B757DF281F75FBEED05349073E9262EF18473DE6D4A84E7C2559589; sso.jd.com=BJ.77c40aaff2ae4b21af1de6c8801bdb61; __jda=91521793.1571023082727881273581.1571023083.1574903895.1574993738.23; __jdc=91521793; ValiCode=YYNJTH7K4PV6; cadata=X3TjTeGIzbq7+loS9Fh4H0Rk2webbgeOQiygDbTeejiXu0RBC7ZuOJGi4I1V2X0c4fSORS3D75OLN9fDjhpZDVUc2bwl7A0KssBOj6YMz+a6ZWJYtwG3Lit0/XUM3f7K; cadataTTL=HSACYD03b3e4+Yndcvs+Kw==; cadataKey=F+08ZMKNfPSgJCMTwX+MiqOnr92scztjueyYU/iq5lgEytRERwWmkE/xafbmHLLA3Ga2BAojRwvQ/CqLq4yUzXlwM5ABc3LkWlMLAGaMIVsCZSVSrPJIUSQLd7+e3d9OsngARBCQlFHMsX0lHiwWOpUovpzx1vNRj0C21imFuHmazf2/kr8crIK2Oh3tFJHh/f/D9jjXioEXpT1cPgmEuse7DKqeW+b3yIY6Whvt1ceg8W4qxn2GHXn6O5bDKvjrMkC34m/bcFG8Kxa+uXwa3FuXHTaPwVNSpuaH4jWK3AtC7FgW2fEua0wUWDJiIJxLOCVzb7Xmsq8i9/1GeVLhCw==; cadataIV=uD30CSRZX8AxCjnMQ/zSW33w/N+udinEaqr3e/YVU07LVKmVEe46hR3K9Fg24/wmFzrH7usmBcidDXP2kdliEaE849cDraIOTKppzSLn5fwKQcEANEGLcq5YaMWGd8L2WhEtFctEW8FzOT77iYUwthfA5Llw9boNjpuFw54seSfrdoR9KNQPivaCPGrtVWRP3WB4ooK98dGrzRAm+bmFSmSp4UlqcHWWMxjBpt/fh2AacC/OGmW1StLG0wcWQ6+bRLoZWPp6qTMiSKloBSQTOOloaq8kaDIrV/SF7WtnijciK7TQmBdXvCnGBhasenoPZA22sxu1kOWF9WHASe3VEg==; cadataSig=lAe4rZIszc+EZ/UhsS/i8j51ERSF4BjfOahWLF8gQUAFAtlgFzoUq6utOzadsq7eQGmJZGJevcTCl/Pv2CJszcrehN9UCWxKthsImCfr2UourQ5Y1DuwLQot7WEQwktfRdexsyc7sLLQS7RuDiWs0fFlFnTgCf2+bOfcJEeRFoRzg0cYjhOioYWHgtinBmPc7zfUr8PX0O2huFSLUNPNi1UEoAG44iPu6IIxJN4o3SNq5R5ts0H2XR9V+rEST2m8g/ukstql7jW4vecWFdTDTRKP3BSw5ELA8Epw+aOvTFRHJ1jMBwOAv9TOMnQzHvahFvs7ylrn27kBCAscw2oG9Q=='}
    response = requests.post(url=url, headers=header, verify=False)
    data1=outputToList(response)

    # 第二部分数据
    url2 = "https://.com/owa/service.svc"
    # 更改获取邮件的数量
    # 查找 %7D%2C%22ViewFilter 部分前面的数据，一般默认数据为25
    # 更改前面的数据为200即可，超过了200也只能获取200
    header2 = {
        'Action': 'FindItem',
        'Cookie': 'X-BackEndCookie=S-1-5-21-1713849901-2797640346-4150151575-1009910=u56Lnp2ejJqBmszJzcrKzMjSyJnIztLLzMac0sfMz8bSz8mZmcmezsnNy8qegYHNz87G0s7N0s3Hq8/OxcvHxc/I; __jdu=1571023082727881273581; pin=zhangvalue; unick=zhangvalue; _tp=cOR1hNRYc3ccyW4KIJwFQQ%3D%3D; _pst=zhangvalue; shshshfpa=b77ba96a-37d4-9aed-2808-8f342d4876af-1571026704; ClientId=A84FA33F1BDC45AD95097F44E4AA1CB9; X-OWA-JS-PSD=1; RoutingKeyCookie=v2:%2fV1xBTCDLP7eJXStxne5F48i2Y5XrC1%2fglQ5zDPnTSQ%3d:d89b6794-bf46-480f-bcf3-4fe09974a4b1@jd.com; unpl=V2_ZzNtbUdTQEZ2W05WexFZDGILRwhKXhMQc19DAHsRXw1mBEJYclRCFX0URlRnGFgUZwEZX0RcQRxFCEdkeB5fA2AFEFlBZxpFK0oYEDlNDEY1WnxZRldAFXEIQlF6KWwGZzMSXHJXRxN1CUVWehBfA2QFG1RCVEUXcQFGZEsebDVXBBtfRVZLJXQ4R2Q5TQAFZQoUXkEaQxFzCEdXeRhVBmEAFFRLV0ATdwxPVEsYbAY%3d; shshshfpb=hK9LDqjkpSdZgURqoxXh1BA%3D%3D; TrackID=1nDRau2MeyWhemIqFXhYVrMilvDeaW_L9GTWtInP3XkUb_pHlROI3gqDILHRvmKnGNWOChdplj7rribkHaGLfbJg5jqHxWzJJJgQFiXmsI1A; __jdv=122132179|ssa.jd.com|-|referral|-|1574240821489; PCSYCityID=CN_110000_110100_110112; shshshfp=ea884eee0c571bd8c126b4842541c636; areaId=1; ipLoc-djd=1-2809-51216-0; 3AB9D23F7A4B3C9B=JZTUZRJV76NL6RWD6IMD2ULQP7WOUEYQ6FF4TIGNXZ2H7USWHQM6CBLV6MHV44GBEQXLJXPDMUQINRMNQZLBHHFDWM; jd.erp.lang=zh_CN; __jda=106621761.1571023082727881273581.1571023083.1574830809.1574903895.22; __jdb=106621761.1.1571023082727881273581|22.1574903895; __jdc=106621761; erp1.jd.com=4BE4D59B25AA1517E56D9E2EF42C5507F13C7B37C69BB9858F71FCB2827045CCE14EF7886578EE467F2CDC456120A8F6BA180E88F0D4536CAEB2DDF805A7AA400E8ED549419AA4106647ACFD1FEB17A4; sso.jd.com=TEST.c86f45124eff4a9b8f5bca663427e5b8; lang=zh-CN; ValiCode=JKJFYQB3GEUV; umail.jd.com=cb43089295b786f1781d9fc614eb7aa670b0dfb5; amail.jd.com=zhangshengfan; cadata=C704nFJMPtLOJ2Pg9Av8P9toOYIXdQEQ6yx7qOncOV+4bDgpJEilmkdRb6U2t+qoCpUCInB4craA3wdjY13d/sXnb8RjgmiQwVNuWIxjvOeTfd+nKqiGElA8E9VcQvNm; cadataTTL=Zcyah4//WFLyJB/GowgzNA==; cadataKey=qyg7V5JHN/IaQ44EQSRV2G38seM7RVyIwNx+ZdGJMdrlN86Nc9sbZMQD2lOSZ0a/GPjDo1Xte91WOewLcS1312bBBxzkRUs1fWlqKXk0pgcB8sxmJb0m4mVt5XGZ259btJSNjNPWkSBXAIO1XCtcXvNY+xennjUvDZFDtdczNk/PDQh20NiOtHus/R5z7zT4QAeNgY/uDl7AXG/JB5JB62OnGZz7gyBlPWA+Q8N8WzlngbGvsmQrDPVvHEbBTIbVDYtAQ9GWf4k6u5y7Mh/8i3Gdi/oPAIzIKxM+36RlLJvEDZd4rIPe4zIzzKTwv7kfgL3vAt+RfxdozqKJyotD+A==; cadataIV=hA/UaANzjIy2fmrNaYuPL2usQc8d4GCyJvLjhNift6q6h98vM5qgRErARaPCaLqNhJs5rAzohDRVWWzrR/4Pl59mXki4oqjAo9DG/dXxUwwFPE3czLSxkQtLWQdeuTmuJw+eZZKLfJuALY+V9Ivl/OJMc02T5jB6qYLS0GFT30bceYpmJ+h+uRx7m8Mr4Esbqa+1tC7tMYTcGX4/OlStHHS5Lq33j/TYh+kEx2JiluNcI0A3XRRaS5i6lyt04n6BBELYX289e3Y31wS2oq7EwsV3JQYKrp44Yl9D6C9McJEY1DPxsxT3dc7w2KQMqe5K07BDNhrYOSC62NtQAVgx/w==; cadataSig=EJQqWRfwUlEylMgnvjFOF1/rlmpYXN+UGOgZK0hSm9ZHcfwy/BS7kJRVwZuAPL7BWkhvR6qpOOEdN+wy5o+bH8bCQ7gNUv8NIc2m26409zTEcrnvesf3cSxf9paEu74lHXGaHikM3pjHgMicbRBUuw3mwPT4pxK2ffXG40J9O6XigyampJnIdwfjYfBrveXZ/+akhRKpSE+nRfxS3ZqOdGHVZ18DNgw+RY0GOwczCrgwp6UPVhU6rw5EddOzuc2itK+nwkpcstpoxVcaY0qhKUHRjDiW+H7pF392CpVpixKDxZKcnM5yBTCROE3e/zG1eNYPgBffZvlDQbhpGlA+7Q==; UC=4167c6142431490cbefd4a759bb8778b; AppcacheVer=15.1.1415.2:zh-cnbase; X-OWA-CANARY=zLyZz4iCO0ihrUH0LTcObCBQMwSlc9cILSDJngKXaXWrZJanEs1x1ufsjMPFMl9OoTvHqfIY_qs.',
        'X-OWA-CANARY': 'zgm6bgPxUkOaeygP6__d8RA1Hzumc9cIKbTglTlHnLvFoOvtxthVA4YnyVAWFg-MAllP6ZRN_ck.',
        'X-OWA-UrlPostData':'%7B%22__type%22%3A%22FindItemJsonRequest%3A%23Exchange%22%2C%22Header%22%3A%7B%22__type%22%3A%22JsonRequestHeaders%3A%23Exchange%22%2C%22RequestServerVersion%22%3A%22Exchange2016%22%2C%22TimeZoneContext%22%3A%7B%22__type%22%3A%22TimeZoneContext%3A%23Exchange%22%2C%22TimeZoneDefinition%22%3A%7B%22__type%22%3A%22TimeZoneDefinitionType%3A%23Exchange%22%2C%22Id%22%3A%22China%20Standard%20Time%22%7D%7D%7D%2C%22Body%22%3A%7B%22__type%22%3A%22FindItemRequest%3A%23Exchange%22%2C%22ItemShape%22%3A%7B%22__type%22%3A%22ItemResponseShape%3A%23Exchange%22%2C%22BaseShape%22%3A%22IdOnly%22%7D%2C%22ParentFolderIds%22%3A%5B%7B%22__type%22%3A%22DistinguishedFolderId%3A%23Exchange%22%2C%22Id%22%3A%22inbox%22%7D%5D%2C%22Traversal%22%3A%22Shallow%22%2C%22Paging%22%3A%7B%22__type%22%3A%22SeekToConditionPageView%3A%23Exchange%22%2C%22BasePoint%22%3A%22Beginning%22%2C%22Condition%22%3A%7B%22__type%22%3A%22RestrictionType%3A%23Exchange%22%2C%22Item%22%3A%7B%22__type%22%3A%22IsEqualTo%3A%23Exchange%22%2C%22Item%22%3A%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22InstanceKey%22%7D%2C%22FieldURIOrConstant%22%3A%7B%22__type%22%3A%22FieldURIOrConstantType%3A%23Exchange%22%2C%22Item%22%3A%7B%22__type%22%3A%22Constant%3A%23Exchange%22%2C%22Value%22%3A%22AQAAAAAAAQwBAAAAB7m%2FbwAAAAA%3D%22%7D%7D%7D%7D%2C%22MaxEntriesReturned%22%3A201%7D%2C%22ViewFilter%22%3A%22All%22%2C%22IsWarmUpSearch%22%3Afalse%2C%22FocusedViewFilter%22%3A-1%2C%22Grouping%22%3Anull%2C%22ShapeName%22%3A%22MailListItem%22%2C%22SortOrder%22%3A%5B%7B%22__type%22%3A%22SortResults%3A%23Exchange%22%2C%22Order%22%3A%22Descending%22%2C%22Path%22%3A%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22ReceivedOrRenewTime%22%7D%7D%2C%7B%22__type%22%3A%22SortResults%3A%23Exchange%22%2C%22Order%22%3A%22Descending%22%2C%22Path%22%3A%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22DateTimeReceived%22%7D%7D%5D%7D%7D'}
      # 'X-OWA-UrlPostData': '%7B%22__type%22%3A%22FindItemJsonRequest%3A%23Exchange%22%2C%22Header%22%3A%7B%22__type%22%3A%22JsonRequestHeaders%3A%23Exchange%22%2C%22RequestServerVersion%22%3A%22Exchange2016%22%2C%22TimeZoneContext%22%3A%7B%22__type%22%3A%22TimeZoneContext%3A%23Exchange%22%2C%22TimeZoneDefinition%22%3A%7B%22__type%22%3A%22TimeZoneDefinitionType%3A%23Exchange%22%2C%22Id%22%3A%22China%20Standard%20Time%22%7D%7D%7D%2C%22Body%22%3A%7B%22__type%22%3A%22FindItemRequest%3A%23Exchange%22%2C%22ItemShape%22%3A%7B%22__type%22%3A%22ItemResponseShape%3A%23Exchange%22%2C%22BaseShape%22%3A%22IdOnly%22%7D%2C%22ParentFolderIds%22%3A%5B%7B%22__type%22%3A%22DistinguishedFolderId%3A%23Exchange%22%2C%22Id%22%3A%22inbox%22%7D%5D%2C%22Traversal%22%3A%22Shallow%22%2C%22Paging%22%3A%7B%22__type%22%3A%22SeekToConditionPageView%3A%23Exchange%22%2C%22BasePoint%22%3A%22Beginning%22%2C%22Condition%22%3A%7B%22__type%22%3A%22RestrictionType%3A%23Exchange%22%2C%22Item%22%3A%7B%22__type%22%3A%22IsEqualTo%3A%23Exchange%22%2C%22Item%22%3A%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22InstanceKey%22%7D%2C%22FieldURIOrConstant%22%3A%7B%22__type%22%3A%22FieldURIOrConstantType%3A%23Exchange%22%2C%22Item%22%3A%7B%22__type%22%3A%22Constant%3A%23Exchange%22%2C%22Value%22%3A%22AQAAAAAAAQwBAAAAB7m%2FNwAAAAA%3D%22%7D%7D%7D%7D%2C%22MaxEntriesReturned%22%3A300%7D%2C%22ViewFilter%22%3A%22All%22%2C%22IsWarmUpSearch%22%3Afalse%2C%22FocusedViewFilter%22%3A-1%2C%22Grouping%22%3Anull%2C%22ShapeName%22%3A%22MailListItem%22%2C%22SortOrder%22%3A%5B%7B%22__type%22%3A%22SortResults%3A%23Exchange%22%2C%22Order%22%3A%22Descending%22%2C%22Path%22%3A%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22ReceivedOrRenewTime%22%7D%7D%2C%7B%22__type%22%3A%22SortResults%3A%23Exchange%22%2C%22Order%22%3A%22Descending%22%2C%22Path%22%3A%7B%22__type%22%3A%22PropertyUri%3A%23Exchange%22%2C%22FieldURI%22%3A%22DateTimeReceived%22%7D%7D%5D%7D%7D'}
    response2 = requests.post(url=url2, headers=header2, verify=False)
    data2 = outputToList(response2)
    data=data1+data2
    # print("***********************合并之前的数据**********************")
    # for i in data:
    #  print(i)
    # 将两次data的数据合并，都是list类型的数据,返回的是一个字典形式
    email_dict=MergeData(data)
    #将数据导入到Excel表格，保存到本地
    ListToExcel(email_dict)
<%@CODEPAGE="65001"%>
<%
    '测试读取 Application 对象中的变量与在页面中定义变量的性能差距
    '作者：孙振国
    '创建时间：2022.03.28
    '

    dim numEach     '循环次数
    dim i           '循环器变量

    numEach = 10000
    

    dim starttime, endtime
    starttime = timer()
    for i = 0 to numEach
        dim c 
        c = application("name")
    next
    endtime = timer()

    dim starttime2, endtime2
    starttime2 = timer()   
    for i = 0 to numEach
        dim c2
        c2 = "/ABCDEFG"
        dim c3
        c3 = c2
    next
    endtime2 = timer()


    
    response.Write"Application:" & FormatNumber((endtime-starttime)*1000,3) & "ms"
    response.Write"<br/>"
    response.Write"Dim:" & FormatNumber((endtime2-starttime2)*1000,3) & "ms"
%>